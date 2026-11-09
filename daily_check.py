#!/usr/bin/env python3
"""Daily Check：讀當天 raw log，在 Daily check 活頁簿中新增一張當天工作表。

做法（跟手動流程一樣）：
  1. 複製範本工作表（例如 1151105）→ 新工作表，名稱為民國日期（例如 1151109），格式完全沿用
  2. 依 raw log 更新有資料的設備列：Memory Usage(E)、CPU Usage(F)、log Analysis(I)
  3. 檢查人員欄 (J)：有給 --replace-name 時，把舊名字換成 --inspector 的名字
  4. F1 區塊中 raw log 沒有的設備，Memory/CPU/log Analysis 填「無資料」
  5. 介面 down 之後沒有 up → 整列標黃；有 up（瞬斷/flapping）不標黃
  6. 另存新檔，不會覆蓋原檔

設備清單直接讀範本工作表，程式裡不含任何設備名稱。

用法一（不用命令提示字元）：
  直接點兩下 daily_check.py（或改名成 daily_check.pyw 就不會跳出黑色視窗），
  依序選「總表」、「當天數據檔」，填檢查人員，就會在總表旁邊存出新檔。
  也可以把總表和數據檔一起拖到 daily_check.py 圖示上。
  第一次執行若沒有 openpyxl，程式會問你要不要自動安裝。

用法二（命令列）：
  pip install openpyxl
  python daily_check.py Daily_check_F1.xlsx 今天的log.log [更多log...] --template 1151105 --inspector 你的名字 --replace-name 舊名字
  （log 也可以給資料夾，會讀裡面所有 .log / .txt；
    也可以直接給網頁工具「Daily Check 匯出」的 .xlsx，會沿用它算好的數值）

常用參數：
  --template 1151105     範本工作表；不給就用最新一天的工作表
  --date 2026-11-09      報表日期；不給就用今天
  --inspector 名字       檢查人員（必填）
  --replace-name 名字    要換掉的舊檢查人員名字（不給就不換）
  --ignore 設備A,設備B    完全忽略的設備
  -o 輸出.xlsx           輸出檔名；不給就用 YYYYMMDD_Daily_check_F1.xlsx
"""
import argparse
import datetime as dt
import os
import re
import sys
from copy import copy

def _ensure_openpyxl():
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        pass
    msg = '需要 openpyxl 套件才能讀寫 Excel，要現在自動安裝嗎？'
    ok = False
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk(); root.withdraw()
        ok = messagebox.askyesno('Daily Check', msg)
        root.destroy()
    except Exception:
        ok = input(msg + ' (y/n) ').strip().lower().startswith('y') if sys.stdin and sys.stdin.isatty() else False
    if not ok:
        sys.exit('需要 openpyxl：請先執行  pip install openpyxl')
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--user', 'openpyxl'])
    import site, importlib
    importlib.reload(site)


_ensure_openpyxl()
import openpyxl  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402

THRESHOLD = 80
YELLOW = 'FFFFFF00'
COL_NAME, COL_VER, COL_MODEL, COL_MEM, COL_CPU, COL_LOG, COL_WHO = 1, 2, 3, 5, 6, 9, 10
LAST_COL = 10

RE_PROMPT = re.compile(r'[\w.\-]+@([A-Za-z0-9][\w.\-]*?)>')
RE_VERSION = re.compile(r'---\s*JUNOS\s+(\S+)|^Junos:\s*(\S+)', re.I)
RE_SYSLOG = re.compile(r'^([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{2}):(\d{2}):(\d{2})\s+([\w.\-]+)\s+(.*)$')
RE_IFACE = re.compile(r'\b((?:xe|ge|et|ae|so|fxp|em|irb|ps|st)-\d+[\d/:.]*)', re.I)
RE_ALARM_LINE = re.compile(r'^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})(?::\d{2})?\s*(?:[A-Z]{2,5})?\s+(?:Major|Minor|Critical)?\s*(.*)$', re.I)
MONTHS = {m: i for i, m in enumerate(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], 1)}


def base_name(h):
    """比對用：去掉 RE0/RE1/node0/node1 後綴、去掉數字前導 0、不分大小寫。"""
    h = re.sub(r'[._\-](re\d|node\d)$', '', str(h or '').strip(), flags=re.I).lower()
    return re.sub(r'(^|[^\d])0+(?=\d)', r'\1', h)


def phys_if(name):
    return re.sub(r'\.\d+$', '', name.lower())


def fmt_dt(t):
    return t.strftime('%m/%d %H:%M')


# ---------------------------------------------------------------- 解析 raw log
class Device:
    def __init__(self, host):
        self.host = host
        self.version = None
        self.res = []          # [(state, mem, cpu)]
        self.alarms = []       # [(ts, text)]
        self.if_events = {}    # iface -> [(time, 'down'|'up', admin_down)]


def parse_logs(paths, year):
    devices = {}

    def dev(h):
        b = base_name(h)
        if b not in devices:
            devices[b] = Device(h)
        return devices[b]

    for path in paths:
        with open(path, encoding='utf-8', errors='replace') as f:
            lines = f.read().splitlines()
        host, mode, buf = None, None, []

        def flush():
            if not mode or not host:
                return
            d = dev(host)
            if mode == 're':
                for blk in re.split(r'^\s*Slot\s+\d+:', '\n'.join(buf), flags=re.M)[1:] or ['\n'.join(buf)]:
                    st = re.search(r'Current state\s+(\S+)', blk)
                    mm = re.search(r'Memory utilization\s+(\d+)\s*percent', blk)
                    idle = re.search(r'CPU utilization:[\s\S]*?Idle\s+(\d+)\s*percent', blk)
                    if mm or idle:
                        d.res.append((st.group(1).lower() if st else '', int(mm.group(1)) if mm else None,
                                      100 - int(idle.group(1)) if idle else None))
            elif mode == 'alarms':
                for l in buf:
                    m = RE_ALARM_LINE.match(l.strip())
                    item = (m.group(1) + ' ' + m.group(2), m.group(3).strip()) if m else None
                    if item and item not in d.alarms:
                        d.alarms.append(item)

        for line in lines:
            pm = RE_PROMPT.search(line)
            if pm:
                flush()
                mode, buf = None, []
                host = pm.group(1)
                cmd = line[pm.end():]
                if re.search(r'show chassis routing-engine', cmd):
                    mode = 're'
                elif re.search(r'show (chassis|system) alarms', cmd):
                    mode = 'alarms'
                dev(host)
                continue
            vm = RE_VERSION.search(line)
            if vm and host:
                dev(host).version = vm.group(1) or vm.group(2)
            if mode:
                if line.strip() == '' and buf:
                    flush()
                    mode, buf = None, []
                else:
                    buf.append(line)
                continue
            m = RE_SYSLOG.match(line)
            if not m:
                continue
            msg = m.group(7)
            im = RE_IFACE.search(msg)
            if not im:
                continue
            arrow = re.search(r'\b(Up|Down)\b\s*->\s*\b(Up|Down)\b', msg, re.I)
            if arrow:
                state = arrow.group(2).lower()
            elif re.search(r'\bdown\b', msg, re.I) or 'LINK_DOWN' in msg:
                state = 'down'
            elif re.search(r'\bup\b', msg, re.I) or 'LINK_UP' in msg:
                state = 'up'
            else:
                continue
            try:
                t = dt.datetime(year, MONTHS[m.group(1)], int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5)))
            except (KeyError, ValueError):
                continue
            admin = state == 'down' and bool(re.search(r'ifAdminStatus\s+down', msg, re.I))
            dev(m.group(6)).if_events.setdefault(phys_if(im.group(1)), []).append((t, state, admin))
        flush()
    return devices


def summarize(d, year):
    """回傳 (mem, cpu, log_analysis_lines, down_list)。"""
    master = [r for r in d.res if r[0] == 'master']
    pool = master or d.res
    mems = [r[1] for r in pool if r[1] is not None]
    cpus = [r[2] for r in pool if r[2] is not None]
    mem = round(max(mems)) if mems else None
    cpu = round(max(cpus)) if cpus else None
    items, downs = [], []
    if mem is not None and mem > THRESHOLD:
        items.append('RE Memory %d%%' % mem)
    if cpu is not None and cpu > THRESHOLD:
        items.append('RE CPU %d%%' % cpu)
    for ts, txt in d.alarms:
        if ts[:4] == str(year):          # 舊年度的 alarm 屬於長期存在，不列（SOP：看有沒有今年的）
            items.append('Alarm: %s (%s)' % (txt, ts))
    for ifn in sorted(d.if_events):
        evs = sorted(d.if_events[ifn])
        downs_n = sum(1 for e in evs if e[1] == 'down')
        if not downs_n:
            continue
        last_state = evs[-1][1]
        if last_state == 'down':
            last_down = evs[-1][0]
            if evs[-1][2]:
                items.append('%s admin down (%s)' % (ifn, fmt_dt(last_down)))
            else:
                items.append('%s down (%s)' % (ifn, fmt_dt(last_down)))
                downs.append((ifn, last_down))
        elif downs_n >= 3:
            first = next(e[0] for e in evs if e[1] == 'down')
            last = evs[-1][0]
            end = last.strftime('%H:%M') if last.date() == first.date() else fmt_dt(last)
            items.append('%s flapping (%s~%s)' % (ifn, fmt_dt(first), end))
        else:
            items.append('%s\t瞬斷，已恢復' % ifn)
    return mem, cpu, items, downs


class Exported:
    """網頁工具「Daily Check 匯出」的 Excel：直接沿用它算好的數值。"""
    def __init__(self, host, version, summary):
        self.host, self.version, self.summary = host, version, summary


def read_export(path, year):
    wb = openpyxl.load_workbook(path, read_only=False)
    ws = wb.worksheets[0]
    out = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        def pct(v):
            if isinstance(v, (int, float)):
                return round(v * 100) if v <= 1.5 else round(v)
            m = re.match(r'\s*(\d+(?:\.\d+)?)\s*%?\s*$', str(v or ''))
            return round(float(m.group(1))) if m else None
        mem, cpu = pct(ws.cell(r, 4).value), pct(ws.cell(r, 5).value)
        log = str(ws.cell(r, 6).value or '')
        if log == '無資料':
            continue
        items = [l for l in log.split('\n') if l and l != 'Normal']
        # 舊年度的 alarm 不列
        items = [l for l in items if not re.match(r'Alarm:.*\((\d{4})-', l) or re.match(r'Alarm:.*\((\d{4})-', l).group(1) == str(year)]
        yellow = (ws.cell(r, 1).fill.fgColor.rgb or '') == YELLOW and ws.cell(r, 1).fill.fill_type == 'solid'
        downs = []
        for l in items:
            m = re.match(r'^(\S+) down \((\d{2}/\d{2} \d{2}:\d{2})\)$', l)
            if m:
                downs.append((m.group(1), dt.datetime.strptime('%d/%s' % (year, m.group(2)), '%Y/%m/%d %H:%M')))
        if yellow and not downs:
            downs.append(('(見 log Analysis)', dt.datetime(year, 1, 1)))
        out[base_name(name)] = Exported(str(name), str(ws.cell(r, 2).value or '') or None, (mem, cpu, items, downs))
    return out


# ---------------------------------------------------------------- 寫入 Excel
def roc(d):
    return '%d%02d%02d' % (d.year - 1911, d.month, d.day)


class SheetExists(RuntimeError):
    pass


def run(workbook, logs, template=None, date=None, inspector=None, replace_name=None, ignore='', output=None, overwrite=False):
    """產生當天工作表，回傳 (輸出檔路徑, 報告文字)。發生問題時丟出 RuntimeError。"""
    class A: pass
    a = A()
    a.workbook, a.logs, a.template, a.date = workbook, logs, template, date
    a.inspector, a.replace_name, a.ignore, a.output = inspector, replace_name, ignore or '', output
    day = dt.date.fromisoformat(a.date) if a.date else dt.date.today()
    paths = []
    for p in a.logs:
        if os.path.isdir(p):
            paths += sorted(os.path.join(p, f) for f in os.listdir(p) if f.lower().endswith(('.log', '.txt')))
        else:
            paths.append(p)
    if not paths:
        raise RuntimeError('找不到 log 檔')

    wb = openpyxl.load_workbook(a.workbook)
    new_name = roc(day)
    if new_name in wb.sheetnames:
        if not overwrite:
            raise SheetExists('工作表 %s 已經存在，請先刪掉或用 --date 指定別天' % new_name)
        wb[new_name].title = '__old_' + new_name   # 同一天重跑：用舊的當天工作表當範本，做完再刪掉（原檔不受影響）
    old_same_day = ('__old_' + new_name) if ('__old_' + new_name) in wb.sheetnames else None
    days = [n for n in wb.sheetnames if re.fullmatch(r'\d{7}', n)]
    tpl_name = a.template or old_same_day or (days[-1] if days else None)
    if tpl_name not in wb.sheetnames:
        raise RuntimeError('找不到範本工作表 %s（可用的：%s）' % (tpl_name, ', '.join(days[-5:])))
    if a.replace_name == 'auto':             # 自動：把範本 F1 第一台設備的檢查人員換成新的檢查人員
        a.replace_name = None
        f1 = False
        for r in range(1, wb[tpl_name].max_row + 1):
            nm, q = wb[tpl_name].cell(r, COL_NAME).value, str(wb[tpl_name].cell(r, COL_VER).value or '')
            if nm and q.upper().startswith('QTY'):
                f1 = str(nm).strip().upper() == 'F1'
            elif f1 and nm:
                a.replace_name = wb[tpl_name].cell(r, COL_WHO).value
                break

    tpl = wb[tpl_name]
    ws = wb.copy_worksheet(tpl)
    ws.title = new_name
    ws.sheet_view.tabSelected = False
    # copy_worksheet 不會複製這些，手動補上
    ws.freeze_panes = tpl.freeze_panes
    ws.sheet_format = copy(tpl.sheet_format)
    ws.print_options = copy(tpl.print_options)
    ws.page_margins = copy(tpl.page_margins)
    ws.page_setup.orientation = tpl.page_setup.orientation
    for cf in tpl.conditional_formatting:
        for rule in cf.rules:
            ws.conditional_formatting.add(str(cf.sqref), rule)
    # 新工作表放在最後一張日期工作表後面
    after = days[-1] if days else tpl_name
    wb.move_sheet(ws, offset=wb.sheetnames.index(after) + 1 - wb.sheetnames.index(new_name))

    devices = {}
    xlsx = [p for p in paths if p.lower().endswith('.xlsx')]
    raw = [p for p in paths if not p.lower().endswith('.xlsx')]
    for p in xlsx:
        devices.update(read_export(p, day.year))
    if raw:
        devices.update(parse_logs(raw, day.year))
    ignore = {base_name(x) for x in a.ignore.split(',') if x.strip()}

    # 找出 F1 區塊（F1 那列之後，到下一個「QTY:」區塊為止）
    in_f1, section = False, ''
    rows_by_base, missing, notify, notes = {}, [], [], []
    no_fill = PatternFill(fill_type=None)
    yellow = PatternFill(fill_type='solid', fgColor=YELLOW)
    for r in range(1, ws.max_row + 1):
        name = ws.cell(r, COL_NAME).value
        qty = str(ws.cell(r, COL_VER).value or '')
        if name and qty.upper().startswith('QTY'):
            section = str(name).strip()
            in_f1 = section.upper() == 'F1'
            continue
        if a.replace_name and ws.cell(r, COL_WHO).value == a.replace_name:
            ws.cell(r, COL_WHO).value = a.inspector
        if not name or not section:
            continue
        b = base_name(name)
        rows_by_base[b] = r
        was_yellow = any((ws.cell(r, c).fill.fgColor.rgb or '') == YELLOW for c in range(1, LAST_COL + 1))
        if b in ignore:
            continue
        d = devices.get(b)
        if d is None:
            if in_f1:
                for c in (COL_MEM, COL_CPU, COL_LOG):
                    ws.cell(r, c).value = '無資料'
                missing.append(name)
            continue
        mem, cpu, items, downs = d.summary if isinstance(d, Exported) else summarize(d, day.year)
        ws.cell(r, COL_MEM).value = '%d%%' % mem if mem is not None else '無資料'
        ws.cell(r, COL_CPU).value = '%d%%' % cpu if cpu is not None else '無資料'
        if mem is None or cpu is None:
            notes.append('%s：log 裡沒有 show chassis routing-engine，Memory/CPU 填「無資料」' % name)
        ws.cell(r, COL_LOG).value = '\n'.join(items) if items else 'Normal'
        if d.version and str(ws.cell(r, COL_VER).value) != d.version:
            notes.append('%s 版本 %s → %s' % (name, ws.cell(r, COL_VER).value, d.version))
            ws.cell(r, COL_VER).value = d.version
        for c in range(1, LAST_COL + 1):
            if downs:
                ws.cell(r, c).fill = yellow
            elif was_yellow:
                ws.cell(r, c).fill = no_fill
        for ifn, t in downs:
            notify.append('%s  %s  down %s' % (name, ifn, fmt_dt(t)))

    if not any(str(ws.cell(r, COL_NAME).value or '').strip().upper() == 'F1' for r in range(1, ws.max_row + 1)):
        raise RuntimeError('範本工作表 %s 裡找不到「F1」區塊，選到的檔案可能不是總表。' % tpl_name)

    extra = [d.host for b, d in devices.items() if b not in rows_by_base and b not in ignore]

    out = a.output or os.path.join(os.path.dirname(os.path.abspath(a.workbook)), day.strftime('%Y%m%d') + '_Daily_check_F1.xlsx')
    if not a.output:
        stem, k = out[:-5], 2
        while os.path.exists(out) or os.path.abspath(out) == os.path.abspath(a.workbook):
            out = '%s (%d).xlsx' % (stem, k); k += 1
    if os.path.abspath(out) == os.path.abspath(a.workbook):
        raise RuntimeError('輸出檔不能跟原檔同名，請用 -o 指定別的檔名')
    if old_same_day:
        del wb[old_same_day]                 # 新的當天工作表已經放在前一天工作表後面
        if tpl_name == old_same_day:
            tpl_name = new_name + '（舊版）'
    wb.active = wb.sheetnames.index(new_name)
    wb.save(out)

    rep = []
    rep.append('已新增工作表 %s（範本 %s），存成 %s' % (new_name, tpl_name, out))
    rep.append('讀了 %d 個檔案，取得 %d 台設備的資料' % (len(paths), len(devices)))
    rep.append('\n需通知客戶（未恢復的 down，已標黃）：')
    rep.append('\n'.join('  ' + x for x in notify) or '  無')
    if missing:
        rep.append('\nF1 缺資料（填「無資料」，共 %d 台）：' % len(missing))
        rep.append('  ' + ', '.join(missing))
    if extra:
        rep.append('\n活頁簿中找不到的設備（未寫入）：')
        rep.append('  ' + ', '.join(extra))
    if notes:
        rep.append('\n備註：')
        rep.append('\n'.join('  ' + x for x in notes))
    rep.append('\n其他區塊（WVPN / Allot / A10 / Tapping）中 log 裡沒有的設備，數值沿用範本 %s。' % tpl_name)


    return out, '\n'.join(rep)


CONFIG = os.path.join(os.path.expanduser('~'), '.daily_check.json')


def _load_cfg():
    try:
        import json
        with open(CONFIG, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cfg(cfg):
    try:
        import json
        with open(CONFIG, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False)
    except Exception:
        pass


def _is_master(path):
    """總表：有很多張民國日期（7 位數字）工作表。"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        return sum(1 for n in wb.sheetnames if re.fullmatch(r'\d{7}', n)) >= 2
    except Exception:
        return False


def gui_main(files):
    try:
        import tkinter
        from tkinter import filedialog, messagebox, simpledialog
    except ImportError:
        sys.exit('這台電腦的 Python 沒有 tkinter，無法開視窗；請改用命令列：\n'
                 '  python daily_check.py 總表.xlsx 數據檔 --inspector 名字')
    root = tkinter.Tk(); root.withdraw()
    try:
        cfg = _load_cfg()
        master, data = None, []
        for f in files:
            if f.lower().endswith('.xlsx') and master is None and _is_master(f):
                master = f
            else:
                data.append(f)
        if not master:
            master = filedialog.askopenfilename(title='1/2 選「總表」（Daily check 活頁簿）',
                                                initialdir=cfg.get('dir') or None,
                                                filetypes=[('Excel', '*.xlsx')])
            if not master:
                return
        if not _is_master(master):
            messagebox.showerror('Daily Check', '這個檔案看起來不是總表（裡面沒有每天的工作表）：\n' + master)
            return
        if not data:
            data = list(filedialog.askopenfilenames(title='2/2 選「當天數據」（網頁匯出的 xlsx 或 log，可多選）',
                                                    initialdir=os.path.dirname(master),
                                                    filetypes=[('數據檔', '*.xlsx *.log *.txt'), ('全部', '*.*')]))
            if not data:
                return
        who = simpledialog.askstring('Daily Check', '檢查人員：', initialvalue=cfg.get('inspector', ''), parent=root)
        if not who:
            return
        cfg.update(inspector=who.strip(), dir=os.path.dirname(master)); _save_cfg(cfg)
        root.config(cursor='watch'); root.update()
        try:
            out, report = run(master, data, inspector=who.strip(), replace_name='auto')
        except SheetExists:
            root.config(cursor=''); root.update()
            if not messagebox.askyesno('Daily Check', '總表裡已經有今天的工作表 %s。\n要用這次的數據重做一張嗎？（原檔不會被改）' % roc(dt.date.today())):
                return
            root.config(cursor='watch'); root.update()
            out, report = run(master, data, inspector=who.strip(), replace_name='auto', overwrite=True)
        messagebox.showinfo('Daily Check 完成', report + '\n\n檔案：' + out)
        try:
            os.startfile(os.path.dirname(out))   # Windows：打開輸出資料夾
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror('Daily Check 失敗', str(e))
    finally:
        root.destroy()


def main():
    args = sys.argv[1:]
    # 點兩下（沒有參數）或把檔案拖到圖示上（只有檔案、沒有 -- 參數）→ 視窗模式
    if not args or all(not x.startswith('-') and os.path.exists(x) for x in args):
        return gui_main(args)
    ap = argparse.ArgumentParser(description='Daily Check：依 raw log 產生當天工作表')
    ap.add_argument('workbook')
    ap.add_argument('logs', nargs='+')
    ap.add_argument('--template')
    ap.add_argument('--date')
    ap.add_argument('--inspector', required=True)
    ap.add_argument('--replace-name')
    ap.add_argument('--ignore', default='')
    ap.add_argument('-o', '--output')
    a = ap.parse_args(args)
    try:
        out, report = run(a.workbook, a.logs, a.template, a.date, a.inspector, a.replace_name, a.ignore, a.output)
    except RuntimeError as e:
        sys.exit(str(e))
    print(report)


if __name__ == '__main__':
    main()
