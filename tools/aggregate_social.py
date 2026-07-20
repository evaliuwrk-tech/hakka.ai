# -*- coding: utf-8 -*-
"""
彙總 FB_IG_流量報告_統計圖表.xlsx 的社群數據,輸出三層 JSON:
  monthly : 各平臺逐月合計
  topics  : 各平臺 × 月份 × 貼文主題 合計
  weekly  : 各平臺逐週合計(週一~週日;僅列至最後一個完整週)

用法:
    python aggregate_social.py 輸出檔.json [起始年月]
    例:python aggregate_social.py social.json 202601
"""
import openpyxl, json, sys, datetime, os

XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    "..", "..", "FB_IG_流量報告_統計圖表.xlsx")
OUT = sys.argv[1] if len(sys.argv) > 1 else "social.json"
START = sys.argv[2] if len(sys.argv) > 2 else "202601"
TODAY = datetime.date.today()
END = f"{TODAY.year}{TODAY.month:02d}"
METRICS = ["posts", "likes", "reach", "shares", "comments", "views", "interactions"]

def month_range(start, end):
    y, m = int(start[:4]), int(start[4:])
    ey, em = int(end[:4]), int(end[4:])
    out = []
    while (y, m) <= (ey, em):
        out.append(f"{y}{m:02d}")
        m += 1
        if m > 12: y, m = y + 1, 1
    return out

MONTHS = month_range(START, END)
def week_monday(d):
    return d - datetime.timedelta(days=d.isoweekday() - 1)

def num(v):
    if v is None: return 0
    if isinstance(v, (int, float)): return v
    try: return float(str(v).replace(",", ""))
    except ValueError: return 0

def parse_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    if isinstance(v, str):
        s = v.strip().replace("/", "-").replace(".", "-")
        try: return datetime.datetime.strptime(s, "%Y-%m-%d").date()
        except ValueError: return None
    return None

def blank():
    return {k: 0 for k in METRICS}

def add(a, r, idx, like_col):
    a["posts"] += 1
    a["likes"] += num(r[idx[like_col]])
    a["reach"] += num(r[idx["觸及人數"]])
    a["shares"] += num(r[idx["分享次數"]])
    a["comments"] += num(r[idx["留言數"]])
    a["views"] += num(r[idx["瀏覽次數"]])
    a["interactions"] += num(r[idx["互動次數"]])

def intify(d):
    for k, v in d.items():
        if isinstance(v, float) and v == int(v): d[k] = int(v)
    return d

wb = openpyxl.load_workbook(XLSX, data_only=True)
out = {"months": MONTHS, "monthly": {}, "topics": {}, "weeks": [], "weekly": {}}
week_set = set()
plat_rows = {}

for sheet, key in (("FB", "fb"), ("IG", "ig")):
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).replace("\n", "").replace("​", "").strip() if c else "" for c in rows[0]]
    like_col = "按讚數和心情數" if sheet == "IG" else "按讚數"
    idx = {n: header.index(n) for n in
           ["發佈日期", "貼文主題", like_col, "觸及人數", "分享次數", "留言數", "瀏覽次數", "互動次數"]}
    monthly = {m: blank() for m in MONTHS}
    topics = {m: {} for m in MONTHS}
    weekly = {}
    for r in rows[1:]:
        d = parse_date(r[idx["發佈日期"]])
        if d is None: continue
        ym = f"{d.year}{d.month:02d}"
        if ym in monthly:
            add(monthly[ym], r, idx, like_col)
            topic = str(r[idx["貼文主題"]] or "未分類").strip() or "未分類"
            t = topics[ym].setdefault(topic, blank())
            add(t, r, idx, like_col)
        # 週彙總:週一為週起,僅收最後完整週(含)以前
        wm = week_monday(d)
        if wm + datetime.timedelta(days=6) < TODAY and ym >= START:
            wk = wm.isoformat()
            week_set.add(wk)
            w = weekly.setdefault(wk, blank())
            add(w, r, idx, like_col)
    out["monthly"][key] = {m: intify(v) for m, v in monthly.items()}
    out["topics"][key] = {m: {t: intify(v) for t, v in tv.items()} for m, tv in topics.items()}
    plat_rows[key] = weekly

weeks = sorted(week_set)[-13:]   # 近 13 個完整週
out["weeks"] = weeks
for key in ("fb", "ig"):
    out["weekly"][key] = {w: intify(plat_rows[key].get(w, blank())) for w in weeks}

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=1)
print(f"已輸出 {OUT}:月份 {MONTHS[0]}~{MONTHS[-1]},週次 {weeks[0]}~{weeks[-1]}(共{len(weeks)}週)")
