# -*- coding: utf-8 -*-
"""
关务月报数据大屏 · 一键生成（Excel 原生看板 + 独立 HTML 大屏）
================================================================
从源工作簿各业务子表抓取汇总数据，一次性产出：
  ① <文件名>_数据大屏.xlsx   — 在原工作簿副本的「数据大屏」子表内直接画 KPI 卡片和 7 个原生图表
  ② <文件名>_数据大屏.html   — 深色科技风全屏 HTML 大屏（SVG 绘制，零外部依赖）

依赖: openpyxl  pywin32       pip install openpyxl pywin32
运行: python make_dashboard.py

复用新月份: 只改 CONFIG 区里的 SRC 即可，其余全部自动。
"""

import atexit
import json
import os
import re
import shutil
import sys
import traceback

import openpyxl                      # 读源工作簿结构
import win32com.client as win32      # 驱动 Excel 原生图表


# ============== CONFIG 按需修改 ==============
SRC              = r"关务月报_10月.xlsx"
OUT_XLSX_SUFFIX  = "数据大屏.xlsx"
OUT_HTML_SUFFIX  = "数据大屏.html"
MAIN_MONTH_IDX   = 9                 # 主视角月份 0-based：Jan=0, ..., Oct=9, Nov=10, Dec=11
AMOUNT_SRC       = r""               # 可选：进出口金额独立文件；留空则从 SRC 尝试读 "进出口金额" 子表
SHOW_TOP10       = False              # True=显示 TOP10 项目报关量；False=板块替换为「进出口金额」
AMOUNT_UNIT      = "RMB"              # 金额显示单位（仅标注用，不影响数值）
# ============================================

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "July", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ---------- 1) 取数 ----------
def num(v):
    if isinstance(v, (int, float)): return float(v)
    return 0.0

def month_matrix(ws, first_row, header_col=1):
    rows = []
    for r in range(first_row, ws.max_row + 1):
        name = ws.cell(r, header_col).value
        if not isinstance(name, str): continue
        name = name.strip()
        if not name or name.lower().startswith("total"): continue
        vals = [num(ws.cell(r, 2 + i).value) for i in range(12)]
        if sum(vals) == 0: continue
        rows.append((name, vals))
    return rows

def total_row(ws, label):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().lower() == label:
            return [num(ws.cell(r, 2 + i).value) for i in range(12)]
    return [0.0] * 12

def location_stats(ws):
    locs = {}
    for r in range(1, ws.max_row + 1):
        loc = ws.cell(r, 15).value
        cnt = ws.cell(r, 16).value
        if isinstance(loc, str) and loc.strip() and isinstance(cnt, (int, float)):
            locs[loc.strip()] = locs.get(loc.strip(), 0) + int(cnt)
    return locs

def top(items, idx, n=10):
    data = sorted(((name, v[idx]) for name, v in items), key=lambda x: -x[1])
    return [{"name": k, "value": round(v, 0)} for k, v in data if v > 0][:n]

def mom(cur, prev):
    if prev <= 0: return None
    return round((cur - prev) / prev * 100, 1)

def _extract_amounts(ws):
    """
    在 matrix 布局的金额子表里启发式找出「入仓金额」「出仓金额」两行的 12 个月数值。
    匹配关键字优先级：行名含 "入/进口/inbound/import" → 入仓金额；含 "出/出口/outbound/export" → 出仓金额。
    找不到则退回 total 行或前两行。
    """
    rows = []
    for r in range(1, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if isinstance(name, str):
            rows.append((name.strip(), r))
    if not rows:
        return [0.0] * 12, [0.0] * 12

    def row_vals(r):
        return [num(ws.cell(r, 2 + i).value) for i in range(12)]

    def pick(rows, keywords_any):
        for name, r in rows:
            low = name.lower()
            if any(k.lower() in low for k in keywords_any):
                return name, row_vals(r)
        return None

    in_row = pick(rows, ["入仓金额", "进口金额", "入仓", "进口", "inbound", "import"])
    out_row = pick(rows, ["出仓金额", "出口金额", "出仓", "出口", "outbound", "export"])

    # 还找不到就退而求其次：看 total 行
    if in_row is None or sum(in_row[1]) == 0:
        for name, r in rows:
            if name.lower() == "total":
                in_row = (name, row_vals(r)); break
    if out_row is None or sum(out_row[1]) == 0:
        if in_row is not None:
            # total 只取到一次就给 out 也复用？不，还是返回 0
            out_row = ("", [0.0] * 12)
        else:
            out_row = ("", [0.0] * 12)

    in_vals  = in_row[1]  if in_row  else [0.0] * 12
    out_vals = out_row[1] if out_row else [0.0] * 12
    return in_vals, out_vals


def extract(src_path, main_idx):
    wb = openpyxl.load_workbook(src_path, data_only=True)

    ws_in = wb["入仓报关单份数"]
    ws_out = wb["出仓报关单份数"]
    in_rows  = month_matrix(ws_in,  3)
    out_rows = month_matrix(ws_out, 3)
    in_total  = total_row(ws_in,  "total")
    out_total = total_row(ws_out, "total")

    ws_i = wb["入仓查验数据"]
    ws_o = wb["出仓查验数据"]
    iinsp_total = total_row(ws_i, "total")
    oinsp_total = total_row(ws_o, "total")
    iinsp_locs = location_stats(ws_i)

    ws_coo = wb["COO证书认证"]
    coo_rows = month_matrix(ws_coo, 3)
    coo_total = total_row(ws_coo, "total")

    ws_diff = wb["库存差异"]
    diff_rows = month_matrix(ws_diff, 3)
    diff_total = [sum(v[i] for _, v in diff_rows) for i in range(12)]

    # ---- 进出口金额 ----
    amount_in_total  = [0.0] * 12
    amount_out_total = [0.0] * 12
    amount_source    = None

    # 1) 先查 SRC 里的候选子表
    for amt_sheet_name in ["进出口金额", "金额", "外贸金额", "进出口总金额", "进出口"]:
        if amt_sheet_name in wb.sheetnames:
            ws_amt = wb[amt_sheet_name]
            a_in, a_out = _extract_amounts(ws_amt)
            if sum(a_in) + sum(a_out) > 0:
                amount_in_total, amount_out_total = a_in, a_out
                amount_source = amt_sheet_name
                break

    # 2) SRC 里没有，再看 CONFIG 里的独立金额文件
    if amount_source is None and AMOUNT_SRC and os.path.exists(AMOUNT_SRC):
        try:
            wb_amt = openpyxl.load_workbook(AMOUNT_SRC, data_only=True)
            ws_amt = None
            for candidate in ["进出口金额", "金额", "外贸金额", "进出口总金额", "Sheet1"]:
                if candidate in wb_amt.sheetnames:
                    ws_amt = wb_amt[candidate]; break
            if ws_amt is None and wb_amt.sheetnames:
                ws_amt = wb_amt[wb_amt.sheetnames[0]]
            if ws_amt:
                a_in, a_out = _extract_amounts(ws_amt)
                if sum(a_in) + sum(a_out) > 0:
                    amount_in_total, amount_out_total = a_in, a_out
                    amount_source = AMOUNT_SRC
        except Exception:
            pass

    data = {
        "period": f"2026年 1-{main_idx + 1}月",
        "main_month": MONTHS[main_idx],
        "months": MONTHS[: main_idx + 1],
        "totals": {
            "inbound":        [round(x) for x in in_total[: main_idx + 1]],
            "outbound":       [round(x) for x in out_total[: main_idx + 1]],
            "in_inspection":  [round(x) for x in iinsp_total[: main_idx + 1]],
            "out_inspection": [round(x) for x in oinsp_total[: main_idx + 1]],
            "coo":            [round(x) for x in coo_total[: main_idx + 1]],
            "diff":           [round(x) for x in diff_total[: main_idx + 1]],
            "amount_in":      [round(x) for x in amount_in_total[: main_idx + 1]],
            "amount_out":     [round(x) for x in amount_out_total[: main_idx + 1]],
        },
        "kpi": {
            "inbound":        {"value": round(in_total[main_idx]),   "prev": round(in_total[main_idx - 1]),
                               "mom": mom(in_total[main_idx],   in_total[main_idx - 1]),
                               "ytd": round(sum(in_total[: main_idx + 1]))},
            "outbound":       {"value": round(out_total[main_idx]),  "prev": round(out_total[main_idx - 1]),
                               "mom": mom(out_total[main_idx],  out_total[main_idx - 1]),
                               "ytd": round(sum(out_total[: main_idx + 1]))},
            "in_inspection":  {"value": round(iinsp_total[main_idx]),"prev": round(iinsp_total[main_idx - 1]),
                               "mom": mom(iinsp_total[main_idx], iinsp_total[main_idx - 1]),
                               "ytd": round(sum(iinsp_total[: main_idx + 1])),
                               "rate": round(iinsp_total[main_idx] / in_total[main_idx] * 100, 2)
                                      if in_total[main_idx] else 0},
            "out_inspection": {"value": round(oinsp_total[main_idx]),"prev": round(oinsp_total[main_idx - 1]),
                               "mom": mom(oinsp_total[main_idx], oinsp_total[main_idx - 1]),
                               "ytd": round(sum(oinsp_total[: main_idx + 1])),
                               "rate": round(oinsp_total[main_idx] / out_total[main_idx] * 100, 2)
                                      if out_total[main_idx] else 0},
            "coo":            {"value": round(coo_total[main_idx]),  "prev": round(coo_total[main_idx - 1]),
                               "mom": mom(coo_total[main_idx],  coo_total[main_idx - 1]),
                               "ytd": round(sum(coo_total[: main_idx + 1]))},
            "diff":           {"value": round(diff_total[main_idx]),"prev": round(diff_total[main_idx - 1]),
                               "mom": mom(diff_total[main_idx], diff_total[main_idx - 1]),
                               "ytd": round(sum(diff_total[: main_idx + 1]))},
            "amount_in":      {"value": round(amount_in_total[main_idx]),
                               "prev": round(amount_in_total[main_idx - 1]),
                               "mom": mom(amount_in_total[main_idx], amount_in_total[main_idx - 1]),
                               "ytd": round(sum(amount_in_total[: main_idx + 1]))},
            "amount_out":     {"value": round(amount_out_total[main_idx]),
                               "prev": round(amount_out_total[main_idx - 1]),
                               "mom": mom(amount_out_total[main_idx], amount_out_total[main_idx - 1]),
                               "ytd": round(sum(amount_out_total[: main_idx + 1]))},
        },
        "top": {
            "inbound":  top(in_rows,  main_idx),
            "outbound": top(out_rows, main_idx),
            "diff":     top(diff_rows, main_idx),
            "coo":      [{"name": n, "value": round(v[main_idx])}
                         for n, v in sorted(coo_rows, key=lambda x: -x[1][main_idx])
                         if v[main_idx] > 0],
        },
        "coo_ytd": [{"name": n, "value": round(sum(v[: main_idx + 1]))}
                    for n, v in sorted(coo_rows, key=lambda x: -sum(x[1][: main_idx + 1]))],
        "inspection_rate": {
            "in":  [round(iinsp_total[i] / in_total[i] * 100, 3) if in_total[i] else 0
                    for i in range(main_idx + 1)],
            "out": [round(oinsp_total[i] / out_total[i] * 100, 3) if out_total[i] else 0
                    for i in range(main_idx + 1)],
        },
        "in_inspection_locs": sorted(iinsp_locs.items(), key=lambda x: -x[1]),
        "amount_source": amount_source,
    }
    return data


# ---------- 2) 生成 Excel 原生看板 ----------
def excel_color(r, g, b): return r + g * 256 + b * 65536

def build_excel(data, src_path, out_path):
    """把 out_path 当作最终输出，内部用带时间戳的临时文件规避文件锁"""
    import time as _time
    tmp_path = os.path.join(os.path.dirname(out_path),
                            f"~tmp_{_time.strftime('%H%M%S')}_{os.path.basename(out_path)}")
    # 如果存在残留先删掉
    for p in (tmp_path, out_path):
        try: os.remove(p)
        except Exception: pass
    _build_excel_inner(data, src_path, tmp_path)
    # 最终改名
    try:
        os.replace(tmp_path, out_path)    # 原子替换（win 上允许 overwrite）
    except Exception:
        shutil.copyfile(tmp_path, out_path)
        try: os.remove(tmp_path)
        except Exception: pass


def _build_excel_inner(data, src_path, out_path):
    BG     = excel_color(4, 18, 31)
    PANEL  = excel_color(13, 43, 66)
    INK    = excel_color(232, 244, 253)
    MUTED  = excel_color(127, 163, 191)
    CYAN   = excel_color(55, 200, 232)
    AMBER  = excel_color(245, 169, 59)
    MINT   = excel_color(53, 217, 168)
    RED    = excel_color(255, 111, 129)
    VIOLET = excel_color(143, 134, 255)
    WHITE  = excel_color(255, 255, 255)

    # 复制源文件为输出
    if os.path.exists(out_path):
        try: os.remove(out_path)
        except Exception: pass
    shutil.copyfile(src_path, out_path)

    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False

    def _safe_exit():
        try:
            for i in range(app.Workbooks.Count, 0, -1):
                try: app.Workbooks(i).Close(False)
                except Exception: pass
        except Exception: pass
        try: app.Quit()
        except Exception: pass
    atexit.register(_safe_exit)

    wb = app.Workbooks.Open(out_path, UpdateLinks=0)

    # --- 隐藏数据源子表 ---
    for sh in list(wb.Worksheets):
        if sh.Name == "_大屏数据源": sh.Delete()
    src = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    src.Name = "_大屏数据源"

    months = data["months"]
    src.Range("A1").Value = "月份"
    src.Range("B1").Value = "入仓报关单"
    src.Range("C1").Value = "出仓报关单"
    src.Range("D1").Value = "入仓查验次数"
    src.Range("E1").Value = "出仓查验次数"
    src.Range("F1").Value = "入仓查验率%"
    src.Range("G1").Value = "出仓查验率%"
    src.Range("H1").Value = "入仓金额"
    src.Range("I1").Value = "库存差异"
    src.Range("J1").Value = "出仓金额"
    for i, _ in enumerate(months):
        r = i + 2
        src.Cells(r, 1).Value = months[i]
        src.Cells(r, 2).Value = data["totals"]["inbound"][i]
        src.Cells(r, 3).Value = data["totals"]["outbound"][i]
        src.Cells(r, 4).Value = data["totals"]["in_inspection"][i]
        src.Cells(r, 5).Value = data["totals"]["out_inspection"][i]
        src.Cells(r, 6).Value = data["inspection_rate"]["in"][i]
        src.Cells(r, 7).Value = data["inspection_rate"]["out"][i]
        src.Cells(r, 8).Value = data["totals"]["amount_in"][i]
        src.Cells(r, 9).Value = data["totals"]["diff"][i]
        src.Cells(r, 10).Value = data["totals"]["amount_out"][i]

    # COO：取 10 月 Top5 + 其他
    coo_oct = {x["name"]: x["value"] for x in data["top"]["coo"]}
    top5 = sorted(coo_oct.items(), key=lambda kv: -kv[1])[:5]
    rest_oct = sum(v for k, v in coo_oct.items() if k not in dict((n, 1) for n, _ in top5))
    rows_coo = top5 + [("其他", rest_oct)]
    src.Range("K1").Value = "证书类型"
    src.Range("L1").Value = "10月"
    for i, (name, v_oct) in enumerate(rows_coo):
        src.Cells(i + 2, 11).Value = name
        src.Cells(i + 2, 12).Value = v_oct

    # TOP10（出仓为锚点）—— 按开关决定是否写入
    if SHOW_TOP10:
        in_map = {x["name"]: x["value"] for x in data["top"]["inbound"]}
        src.Range("O1").Value = "项目"
        src.Range("P1").Value = "出仓"
        src.Range("Q1").Value = "入仓"
        for i, item in enumerate(data["top"]["outbound"]):
            src.Cells(i + 2, 15).Value = item["name"]
            src.Cells(i + 2, 16).Value = item["value"]
            src.Cells(i + 2, 17).Value = in_map.get(item["name"], 0)

    # 查验地点
    src.Range("S1").Value = "查验地点"
    src.Range("T1").Value = "次数"
    for i, (name, v) in enumerate(data["in_inspection_locs"]):
        src.Cells(i + 2, 19).Value = name
        src.Cells(i + 2, 20).Value = v

    src.Visible = 0   # xlSheetHidden

    # --- 准备「数据大屏」子表：存在就复用清空，不存在就新建 ---
    ws = None
    for sh in list(wb.Worksheets):
        if sh.Name.strip() == "数据大屏":
            ws = sh; break
    if ws is None:
        # 没找到就新建一个（插到最后）
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "数据大屏"

    ws.Activate
    ws.Cells.UnMerge()
    ws.Cells.Clear()
    try:
        ws.Shapes.SelectAll()
        app.Selection.Delete()
    except Exception:
        pass

    ws.Tab.Color = CYAN
    ws.Range("A1:W58").Interior.Color = BG
    ws.Cells.Font.Name = "Microsoft YaHei"
    ws.Cells.Font.Color = INK
    try: app.ActiveWindow.DisplayGridlines = False
    except Exception: pass

    ws.Columns("A").ColumnWidth = 2.2
    ws.Columns("B:W").ColumnWidth = 15.6
    for r, h in [(1, 8), (2, 34), (3, 20), (4, 8), (5, 20), (6, 40), (7, 17), (8, 17), (9, 10)]:
        ws.Rows(r).RowHeight = h
    ws.Rows("10:58").RowHeight = 14

    # 标题
    ws.Range("B2:W2").Merge()
    ws.Range("B2").Value = "关务月报数据大屏"
    ws.Range("B2").Font.Size = 24
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = WHITE
    ws.Range("B2").HorizontalAlignment = -4131   # xlCenter
    ws.Range("B2").VerticalAlignment = -4108

    ws.Range("B3:W3").Merge()
    ws.Range("B3").Value = (
        f"2026年1-{len(months)}月  ｜  主视角：2026年{data['main_month']}  ｜  "
        "数据来源：入仓/出仓报关单份数、入仓/出仓查验数据、COO证书认证、库存差异 子表"
    )
    ws.Range("B3").Font.Size = 10
    ws.Range("B3").Font.Color = MUTED
    ws.Range("B3").HorizontalAlignment = -4131

    # --- 6 个 KPI 卡片 ---
    kpis = [
        ("入仓报关单份数",   data["kpi"]["inbound"]["value"],        "份", data["kpi"]["inbound"]["mom"],
         f"累计 {data['kpi']['inbound']['ytd']:,} 份",                                   CYAN),
        ("出仓报关单份数",   data["kpi"]["outbound"]["value"],       "份", data["kpi"]["outbound"]["mom"],
         f"累计 {data['kpi']['outbound']['ytd']:,} 份",                                   AMBER),
        ("入仓查验次数",     data["kpi"]["in_inspection"]["value"], "次", data["kpi"]["in_inspection"]["mom"],
         f"查验率 {data['kpi']['in_inspection']['rate']:.2f}%",                           VIOLET),
        ("出仓查验次数",     data["kpi"]["out_inspection"]["value"],"次", data["kpi"]["out_inspection"]["mom"],
         f"查验率 {data['kpi']['out_inspection']['rate']:.2f}%",                          excel_color(111, 143, 240)),
        ("COO证书认证",      data["kpi"]["coo"]["value"],             "份", data["kpi"]["coo"]["mom"],
         f"累计 {data['kpi']['coo']['ytd']:,} 份",                                        MINT),
        ("库存差异",         data["kpi"]["diff"]["value"],                  "", data["kpi"]["diff"]["mom"],
         f"累计 {data['kpi']['diff']['ytd']:,}",                                          RED),
    ]
    start_cols = [3, 6, 9, 12, 15, 18]
    for (label, val, unit, mom, foot, color), c0 in zip(kpis, start_cols):
        c1 = c0 + 2
        rng = ws.Range(ws.Cells(5, c0), ws.Cells(8, c1))
        rng.Interior.Color = PANEL
        rng.HorizontalAlignment = -4131
        rng.VerticalAlignment = -4108

        ws.Range(ws.Cells(5, c0), ws.Cells(5, c1)).Merge()
        ws.Cells(5, c0).Value = label + "  · " + data["main_month"]
        ws.Cells(5, c0).Font.Size = 10.5
        ws.Cells(5, c0).Font.Color = excel_color(157, 192, 216)

        ws.Range(ws.Cells(6, c0), ws.Cells(6, c1)).Merge()
        ws.Cells(6, c0).Value = val
        ws.Cells(6, c0).NumberFormatLocal = "#,##0"
        ws.Cells(6, c0).Font.Size = 28
        ws.Cells(6, c0).Font.Bold = True
        ws.Cells(6, c0).Font.Color = WHITE

        ws.Range(ws.Cells(7, c0), ws.Cells(7, c1)).Merge()
        arrow = "▲" if mom >= 0 else "▼"
        ws.Cells(7, c0).Value = f"{arrow} 环比 {('+' if mom >= 0 else '-')}{abs(mom):.1f}%"
        ws.Cells(7, c0).Font.Size = 10.5
        ws.Cells(7, c0).Font.Bold = True
        ws.Cells(7, c0).Font.Color = MINT if mom >= 0 else RED

        ws.Range(ws.Cells(8, c0), ws.Cells(8, c1)).Merge()
        ws.Cells(8, c0).Value = foot
        ws.Cells(8, c0).Font.Size = 9.5
        ws.Cells(8, c0).Font.Color = MUTED

        left_edge = ws.Range(ws.Cells(5, c0), ws.Cells(8, c0))
        left_edge.Borders(7).LineStyle = 1
        left_edge.Borders(7).Weight   = 4
        left_edge.Borders(7).Color    = color

    # --- 7 个原生图表 ---
    X0 = ws.Cells(11, 3).Left
    Y1 = ws.Cells(11, 3).Top
    Y2 = ws.Cells(28, 3).Top
    Y3 = ws.Cells(43, 3).Top

    def add_chart(x, y, w, h, ctype):
        return ws.ChartObjects().Add(x, y, w, h).Chart

    def style_chart(ch, title):
        try:
            ch.HasTitle = True
            ch.ChartTitle.Text = title
            ch.ChartTitle.Font.Size = 12
            ch.ChartTitle.Font.Bold = True
            ch.ChartTitle.Font.Color = INK
            ch.ChartTitle.Font.Name = "Microsoft YaHei"
        except Exception: pass
        try:
            dark = excel_color(10, 35, 56)
            ch.ChartArea.Format.Fill.ForeColor.RGB = dark
            ch.ChartArea.Format.Line.Visible = 0
            ch.PlotArea.Format.Fill.ForeColor.RGB = dark
            ch.PlotArea.Format.Line.Visible = 0
            ch.ChartArea.Font.Name = "Microsoft YaHei"
            ch.ChartArea.Font.Size = 9
            ch.ChartArea.Font.Color = MUTED
        except Exception: pass
        try:
            for ax in (1, 2):
                a = ch.Axes(ax)
                a.TickLabels.Font.Size = 8.5
                a.TickLabels.Font.Color = MUTED
                a.HasMajorGridlines = (ax == 2)
                a.MajorGridlines.Format.Line.ForeColor.RGB = excel_color(35, 70, 100)
        except Exception: pass

    def series_color(ch, idx, color, line=False):
        s = ch.SeriesCollection(idx)
        if line:
            s.Format.Line.ForeColor.RGB = color
            s.Format.Line.Weight = 2.2
        else:
            try: s.Format.Fill.ForeColor.RGB = color
            except Exception: pass

    def set_series(ch, specs):
        sc = ch.SeriesCollection()
        while sc.Count > len(specs): sc(sc.Count).Delete()
        for i, (nm, cats, vals) in enumerate(specs, start=1):
            formula = f'=SERIES("{nm}",{cats},{vals},{i})'
            if i <= sc.Count: sc(i).Formula = formula
            else: sc.NewSeries().Formula = formula

    # 1 进出仓趋势
    ch = add_chart(X0, Y1, 1100, 240, 65)   # xlLineMarkersStacked
    ch.SetSourceData(Source=src.Range("A1:C11"), PlotBy=2)
    style_chart(ch, "进出仓报关单份数趋势（份）")
    series_color(ch, 1, AMBER, line=True)
    series_color(ch, 2, CYAN,  line=True)

    # 2 查验次数
    ch = add_chart(X0 + 1120, Y1, 620, 240, 51)
    ch.SetSourceData(Source=src.Range("A1:C11"), PlotBy=2)
    style_chart(ch, "进出仓查验次数（次）")
    series_color(ch, 1, VIOLET)
    series_color(ch, 2, excel_color(79, 127, 168))

    # 3 库存差异
    ch = add_chart(X0, Y2, 570, 200, 51)
    ch.SetSourceData(Source=src.Range("A1:A11,I1:I11"), PlotBy=2)
    style_chart(ch, "库存差异月度走势")
    series_color(ch, 1, RED)

    # 4 COO 环形
    ch = add_chart(X0 + 582, Y2, 560, 200, -4120)   # xlDoughnut
    set_series(ch, [("10月认证", "'_大屏数据源'!$K$2:$K$7",
                                  "'_大屏数据源'!$L$2:$L$7")])
    style_chart(ch, "COO证书认证构成（10月）")
    ch.ChartGroups(1).DoughnutHoleSize = 52
    ch.HasLegend = True
    ch.Legend.Position = -4107   # xlLegendPositionBottom
    ch.Legend.Font.Size = 8
    ch.Legend.Font.Color = MUTED
    coo_colors = [excel_color(53,217,168), excel_color(79,214,192),
                  excel_color(63,182,216), excel_color(111,143,240),
                  excel_color(176,139,255), excel_color(91,117,144)]
    ser = ch.SeriesCollection(1)
    for i in range(1, 7):
        ser.Points(i).Format.Fill.ForeColor.RGB = coo_colors[i - 1]
        ser.Points(i).Format.Line.ForeColor.RGB = excel_color(8, 36, 58)

    # 5 进出口金额（簇状柱） —— 原 TOP10 占位；数据为空时跳过
    has_amount = (sum(data["totals"]["amount_in"]) + sum(data["totals"]["amount_out"]) > 0)
    if SHOW_TOP10 and data.get("top", {}).get("outbound"):
        # 保留旧的 TOP10 图逻辑（按开关）
        ch = add_chart(X0 + 1154, Y2, 585, 200, 57)   # xlBarClustered
        set_series(ch, [("出仓报关单", "'_大屏数据源'!$O$2:$O$11",
                                      "'_大屏数据源'!$P$2:$P$11"),
                        ("入仓报关单", "'_大屏数据源'!$O$2:$O$11",
                                      "'_大屏数据源'!$Q$2:$Q$11")])
        style_chart(ch, "项目TOP10报关量（10月）")
        series_color(ch, 1, AMBER)
        series_color(ch, 2, CYAN)
    elif has_amount:
        # 新：进出口金额趋势（簇状柱 + 次轴折线可选）
        ch = add_chart(X0 + 1154, Y2, 585, 200, 51)   # xlColumnClustered
        set_series(ch, [("入仓金额", "'_大屏数据源'!$A$2:$A$11",
                                    "'_大屏数据源'!$H$2:$H$11"),
                        ("出仓金额", "'_大屏数据源'!$A$2:$A$11",
                                    "'_大屏数据源'!$J$2:$J$11")])
        style_chart(ch, f"进出口金额趋势（{AMOUNT_UNIT}）")
        series_color(ch, 1, CYAN)
        series_color(ch, 2, AMBER)
        # 金额数据用千分位
        try:
            for idx in (1, 2):
                ch.SeriesCollection(idx).DataLabels.ShowValue = False
        except Exception: pass
    # else：没有金额也不显示 TOP10 → 空槽位（用户自行决定）

    # 6 查验地点
    ch = add_chart(X0, Y3, 570, 170, 57)
    ch.SetSourceData(Source=src.Range("S1:T7"), PlotBy=2)
    style_chart(ch, "入仓查验地点分布（累计）")
    series_color(ch, 1, excel_color(111, 208, 255))

    # 7 查验率趋势 —— 扩展为入+出 双线
    ch = add_chart(X0 + 582, Y3, 570, 170, 65)
    set_series(ch, [("入仓查验率", "'_大屏数据源'!$A$2:$A$11",
                                    "'_大屏数据源'!$F$2:$F$11"),
                    ("出仓查验率", "'_大屏数据源'!$A$2:$A$11",
                                    "'_大屏数据源'!$G$2:$G$11")])
    style_chart(ch, "进出仓查验率趋势（%）")
    series_color(ch, 1, MINT, line=True)
    series_color(ch, 2, CYAN, line=True)

    # 脚注
    ws.Range("B57:W57").Merge()
    ws.Range("B57").Value = (
        "口径：查验率 = 查验次数 ÷ 报关单份数；环比为 10 月较 9 月；"
        "累计为 2026 年 1-10 月合计。原始明细见各业务子表，“_大屏数据源” 为图表取数区（已隐藏）。"
    )
    ws.Range("B57").Font.Size = 9
    ws.Range("B57").Font.Color = excel_color(95, 127, 151)

    try: app.ActiveWindow.Zoom = 60
    except Exception: pass
    ws.Range("B2").Select()

    app.CutCopyMode = False

    # Excel COM 进程在受限环境下可能无法 SaveAs 到任意路径，
    # 先保存到用户临时目录，再由 Python 复制到目标位置（Python 文件操作不受限）
    import datetime as _dt
    import tempfile
    tmp_dir = tempfile.gettempdir()
    stamp = _dt.datetime.now().strftime("%H%M%S") + str(_dt.datetime.now().microsecond // 1000)
    tmp_path = os.path.join(tmp_dir, f"gwdt_{stamp}.xlsx")
    try:
        wb.SaveAs(tmp_path, FileFormat=51)
    except Exception:
        # 带 UNC\\localhost\C$ 的绝对路径也不行时，退而用 xls 格式
        tmp_path = os.path.join(tmp_dir,
                                f"gwdt_{_time.strftime('%H%M%S%f')}.xls")
        wb.SaveAs(tmp_path, FileFormat=56)

    try: wb.Close(False)
    except Exception: pass
    try: app.Quit()
    except Exception: pass

    if os.path.exists(tmp_path):
        # 把刚才的 SaveAs 输出复制到真正的目标位置
        try:
            os.replace(tmp_path, out_path)
        except Exception:
            shutil.copyfile(tmp_path, out_path)
            try: os.remove(tmp_path)
            except Exception: pass
    else:
        raise RuntimeError(f"SaveAs 未生成临时文件: {tmp_path}")

    try: atexit.unregister(_safe_exit)
    except Exception: pass


# ---------- 3) 生成 HTML 大屏 ----------
HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>关务月报数据大屏 · {{PERIOD}}</title>
<style>
  :root{
    --bg:#04121f; --panel:#0a2338; --panel-2:#0d2c45;
    --line:rgba(89,166,214,.18); --ink:#e8f4fd; --muted:#7fa3bf;
    --cyan:#37c8e8; --cyan-dim:rgba(55,200,232,.14);
    --amber:#f5a93b; --mint:#35d9a8; --red:#ff6f81; --violet:#8f86ff;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  html{height:100%}
  body{
    background:
      radial-gradient(1200px 620px at 12% -10%, rgba(55,200,232,.16), transparent 60%),
      radial-gradient(1000px 520px at 92% 0%, rgba(143,134,255,.13), transparent 62%),
      linear-gradient(180deg,#04121f 0%,#061a2b 55%,#04121f 100%);
    background-attachment: fixed;
    color:var(--ink);
    font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif;
    overflow-x:hidden;
    overflow-y:auto;
  }
  /* 深色滚动条美化 */
  body::-webkit-scrollbar{width:8px}
  body::-webkit-scrollbar-track{background:rgba(55,200,232,.04)}
  body::-webkit-scrollbar-thumb{background:rgba(55,200,232,.35);border-radius:4px}
  body::-webkit-scrollbar-thumb:hover{background:rgba(55,200,232,.55)}
  .screen{min-height:100vh;padding:16px 20px 40px;display:flex;flex-direction:column;gap:14px}
  header{display:flex;align-items:center;justify-content:space-between;gap:20px;height:56px;flex:0 0 auto}
  .brand{display:flex;align-items:center;gap:14px}
  .mark{width:44px;height:44px;border-radius:12px;position:relative;flex:0 0 auto;
        background:linear-gradient(150deg,#37c8e8,#1d7fb5);box-shadow:0 0 22px rgba(55,200,232,.45)}
  .mark::after{content:"";position:absolute;inset:10px;border-radius:6px;
               border:2px solid rgba(4,18,31,.72);border-bottom-color:transparent;border-right-color:transparent;transform:rotate(-8deg)}
  h1{font-size:23px;letter-spacing:3px;font-weight:700;line-height:1.2}
  .sub{font-size:11.5px;color:var(--muted);letter-spacing:1.6px;margin-top:4px}
  .head-right{display:flex;align-items:center;gap:10px}
  .chip{font-size:12px;padding:6px 12px;border-radius:999px;color:#bfe6f6;
        border:1px solid rgba(55,200,232,.32);background:rgba(55,200,232,.08);letter-spacing:.6px}
  .chip b{color:#fff;font-weight:600}
  #clock{font-variant-numeric:tabular-nums;letter-spacing:1px;font-size:13px;color:#cfe8f7}
  .kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;flex:0 0 auto}
  .kpi{position:relative;padding:13px 15px 12px;border-radius:12px;overflow:hidden;
       background:linear-gradient(160deg,rgba(20,60,92,.92),rgba(9,32,51,.94));
       border:1px solid var(--line);box-shadow:inset 0 1px 0 rgba(255,255,255,.04),0 8px 22px rgba(0,0,0,.28)}
  .kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent,var(--cyan));box-shadow:0 0 16px var(--accent,var(--cyan))}
  .kpi .k-label{font-size:12.5px;color:#9dc0d8;letter-spacing:.6px;display:flex;justify-content:space-between;align-items:center;gap:8px;white-space:nowrap}
  .kpi .k-label>span:first-child{overflow:hidden;text-overflow:ellipsis}
  .kpi .k-tag{font-size:10.5px;color:var(--muted);border:1px solid rgba(127,163,191,.28);border-radius:6px;padding:1px 6px;flex:0 0 auto}
  .kpi .k-val{font-size:34px;font-weight:700;letter-spacing:.5px;line-height:1.15;margin-top:6px;font-variant-numeric:tabular-nums;color:#fff;text-shadow:0 0 18px rgba(55,200,232,.28)}
  .kpi .k-val small{font-size:13px;font-weight:500;color:#9dc0d8;margin-left:4px;letter-spacing:0}
  .kpi .k-foot{display:flex;align-items:center;gap:6px;margin-top:7px;font-size:11px;color:#8fb2ca;white-space:normal;flex-wrap:wrap;row-gap:2px;overflow:visible}
  .delta{font-weight:700;font-variant-numeric:tabular-nums}
  .delta.up{color:var(--mint)}.delta.down{color:var(--red)}
  .k-foot .sep{width:1px;height:10px;background:rgba(127,163,191,.32)}
  main{flex:1 1 auto;display:grid;grid-template-columns:repeat(12,1fr);grid-template-rows:auto auto auto;gap:14px}
  .card{grid-column:span 4;grid-row:span 1;min-height:220px;display:flex;flex-direction:column;
        padding:11px 14px 8px;border-radius:12px;
        background:linear-gradient(165deg,rgba(13,44,69,.92),rgba(7,26,42,.94));
        border:1px solid var(--line);box-shadow:0 10px 26px rgba(0,0,0,.26)}
  .card-hd{display:flex;align-items:center;justify-content:space-between;flex:0 0 auto;padding-bottom:6px;border-bottom:1px solid rgba(89,166,214,.12)}
  .card-hd h2{font-size:13.5px;font-weight:600;letter-spacing:1.2px;color:#dceefb;display:flex;align-items:center;gap:8px}
  .card-hd h2::before{content:"";width:4px;height:13px;border-radius:2px;background:var(--accent,var(--cyan));box-shadow:0 0 10px var(--accent,var(--cyan))}
  .card-hd .hint{font-size:10.5px;color:var(--muted);letter-spacing:.5px}
  .chart{flex:1 1 auto;min-height:180px;position:relative;display:flex}
  .chart>svg{width:100%;height:100%}
  .coo-body{flex:1 1 auto;min-height:200px;display:flex;align-items:center;gap:6px}
  .coo-body .donut{flex:0 0 46%;height:100%;min-width:0}
  .coo-legend{flex:1 1 auto;min-width:0;display:flex;flex-direction:column;justify-content:center;gap:2px}
  .coo-legend .row{display:flex;align-items:baseline;gap:6px;font-size:11px;color:#cfe6f6;line-height:1.35}
  .coo-legend .dot{width:8px;height:8px;border-radius:2px;flex:0 0 auto;align-self:center}
  .coo-legend .nm{flex:1 1 auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .coo-legend .vl{font-weight:700;color:#fff;font-variant-numeric:tabular-nums}
  .coo-legend .pc{color:#7fa3bf;font-size:10px;width:32px;text-align:right;font-variant-numeric:tabular-nums}
  .trend{grid-column:span 7;grid-row:span 2}
  .inspect{grid-column:span 5;grid-row:span 1}
  .diff{grid-column:span 5;grid-row:span 1}
  .coo{grid-column:span 4}
  .top{grid-column:span 4}
  .amount{grid-column:span 4}
  .loc{grid-column:span 4}
  svg{display:block;width:100%;height:100%}
  footer{flex:0 0 auto;display:flex;justify-content:space-between;font-size:10.5px;color:#5f7f97;letter-spacing:.4px}
  @media (max-width:1440px){h1{font-size:20px}.kpi .k-val{font-size:28px}}
</style>
</head>
<body>
<div class="screen">
  <header>
    <div class="brand">
      <div class="mark"></div>
      <div>
        <h1>关务月报数据大屏</h1>
        <div class="sub">CUSTOMS MONTHLY REPORT · MAERSK QIANHAI</div>
      </div>
    </div>
    <div class="head-right">
      <span class="chip">主视角 <b>{{MAIN_LABEL}}</b></span>
      <span class="chip">趋势区间 <b>{{PERIOD}}</b></span>
      <span id="clock"></span>
    </div>
  </header>

  <section class="kpis" id="kpis"></section>

  <main>
    <section class="card trend" style="--accent:var(--cyan)">
      <div class="card-hd"><h2>进出仓报关单份数趋势</h2><span class="hint">单位：份 · {{PERIOD}}</span></div>
      <div class="chart" id="ch-trend"></div>
    </section>
    <section class="card inspect" style="--accent:var(--violet)">
      <div class="card-hd"><h2>进出仓查验监控</h2><span class="hint">柱：查验次数 · 线：查验率</span></div>
      <div class="chart" id="ch-inspect"></div>
    </section>
    <section class="card diff" style="--accent:var(--red)">
      <div class="card-hd"><h2>库存差异月度走势</h2><span class="hint">各项目差异数量合计</span></div>
      <div class="chart" id="ch-diff"></div>
    </section>
    <section class="card coo" style="--accent:var(--mint)">
      <div class="card-hd"><h2>COO 证书认证构成</h2><span class="hint">{{MAIN_LABEL}}</span></div>
      <div class="coo-body">
        <div class="donut" id="ch-coo"></div>
        <div class="coo-legend" id="coo-legend"></div>
      </div>
    </section>
    <section class="card amount" style="--accent:var(--amber)">
      <div class="card-hd"><h2>进出口金额月度走势</h2><span class="hint">单位：{{AMOUNT_UNIT}} · {{PERIOD}}</span></div>
      <div class="chart" id="ch-amount"></div>
    </section>
    <section class="card loc" style="--accent:#6fd0ff">
      <div class="card-hd"><h2>入仓查验地点分布</h2><span class="hint">累计查验次数</span></div>
      <div class="chart" id="ch-loc"></div>
    </section>
  </main>

  <footer>
    <span>数据来源：{{SRC_BASENAME}} — 入仓/出仓报关单份数、入仓/出仓查验数据、COO证书认证、库存差异{{AMOUNT_SRC_LABEL}}</span>
    <span>口径说明：查验率 = 查验次数 ÷ 报关单份数（入仓 / 出仓分别统计）；环比 = 较上月；累计 = {{PERIOD}} 合计</span>
  </footer>
</div>

<script>
const DATA = {{DATA_JSON}};

const C = {cyan:"#37c8e8",amber:"#f5a93b",mint:"#35d9a8",red:"#ff6f81",violet:"#8f86ff",grid:"rgba(89,166,214,.16)",axis:"#7fa3bf"};
const fmt = n => (Math.round(n)).toLocaleString("en-US");
const NS = "http://www.w3.org/2000/svg";
const COO_COLORS = ["#35d9a8","#4fd6c0","#3fb6d8","#6f8ff0","#b08bff","#5b7590"];

function svgEl(w,h){const s=document.createElementNS(NS,"svg");s.setAttribute("viewBox",`0 0 ${w} ${h}`);s.setAttribute("preserveAspectRatio","none");s.style.width="100%";s.style.height="100%";return s}
function mk(tag,attrs){const e=document.createElementNS(NS,tag);for(const k in attrs)e.setAttribute(k,attrs[k]);return e}
function niceMax(v,steps=4){if(v<=0)return 10;const raw=v/steps,mag=Math.pow(10,Math.floor(Math.log10(raw)));const norm=raw/mag;const mult=norm<=1?1:norm<=2?2:norm<=2.5?2.5:norm<=5?5:10;return mult*mag*steps}
function text(x,y,str,opt={}){const t=mk("text",{x,y,fill:opt.fill||C.axis,"font-size":opt.size||11,"text-anchor":opt.anchor||"start","font-weight":opt.weight||400,"font-family":'"Microsoft YaHei","PingFang SC",sans-serif',"dominant-baseline":opt.baseline||"auto"});if(opt.opacity)t.setAttribute("opacity",opt.opacity);t.textContent=str;return t}
function gridY(svg,p,x0,x1,max,steps,fmtFn){for(let i=0;i<=steps;i++){const y=p.t+(p.h)*(1-i/steps);svg.appendChild(mk("line",{x1:x0,y1:y,x2:x1,y2:y,stroke:C.grid,"stroke-width":1,"stroke-dasharray":i?"3 4":"0"}));svg.appendChild(text(x0-8,y+3.5,fmtFn(max*i/steps),{anchor:"end",size:10,fill:"#6f92ab"}))}}
function smoothPath(pts){if(pts.length<2)return"";let d=`M${pts[0][0]},${pts[0][1]}`;for(let i=0;i<pts.length-1;i++){const p0=pts[i-1]||pts[i],p1=pts[i],p2=pts[i+1],p3=pts[i+2]||p2;const c1x=p1[0]+(p2[0]-p0[0])/6,c1y=p1[1]+(p2[1]-p0[1])/6;const c2x=p2[0]-(p3[0]-p1[0])/6,c2y=p2[1]-(p3[1]-p1[1])/6;d+=` C${c1x},${c1y} ${c2x},${c2y} ${p2[0]},${p2[1]}`}return d}

function renderTrend(box){
  box.innerHTML="";const W=box.clientWidth,H=box.clientHeight;if(!W||!H)return;
  const p={t:26,r:20,b:26,l:56,h:H-52};const s=svgEl(W,H);
  const lbs=DATA.months,n=lbs.length,iw=W-p.l-p.r;
  const max=niceMax(Math.max(...DATA.totals.outbound),4);gridY(s,p,p.l,W-p.r,max,4,fmt);
  const X=i=>p.l+iw*i/(n-1),Y=v=>p.t+p.h*(1-v/max);
  lbs.forEach((l,i)=>s.appendChild(text(X(i),H-8,l,{anchor:"middle",size:10.5})));
  [["outbound","出仓报关单",C.amber],["inbound","入仓报关单",C.cyan]].forEach(([k,name,color])=>{
    const pts=DATA.totals[k].map((v,i)=>[X(i),Y(v)]);const gid="g"+k;
    const defs=mk("defs",{});const lg=mk("linearGradient",{id:gid,x1:0,y1:0,x2:0,y2:1});
    lg.appendChild(mk("stop",{offset:"0%","stop-color":color,"stop-opacity":.28}));
    lg.appendChild(mk("stop",{offset:"100%","stop-color":color,"stop-opacity":0}));
    defs.appendChild(lg);s.appendChild(defs);const line=smoothPath(pts);
    s.appendChild(mk("path",{d:`${line} L${X(n-1)},${p.t+p.h} L${X(0)},${p.t+p.h} Z`,fill:`url(#${gid})`,stroke:"none"}));
    s.appendChild(mk("path",{d:line,fill:"none",stroke:color,"stroke-width":2.4,"stroke-linecap":"round"}));
    pts.forEach((pt,i)=>{const c=mk("circle",{cx:pt[0],cy:pt[1],r:i===n-1?4.2:2.8,fill:"#04121f",stroke:color,"stroke-width":2});c.appendChild(mk("title",{})).textContent=lbs[i]+" "+name+"："+fmt(DATA.totals[k][i]);s.appendChild(c)});
    const last=pts[n-1];s.appendChild(mk("circle",{cx:last[0],cy:last[1],r:9,fill:color,opacity:.16}));
    const lbl=fmt(DATA.totals[k][n-1]);const ly=last[1]+(k==="outbound"?-11:17);
    s.appendChild(text(W-4,ly,lbl,{anchor:"end",size:11.5,fill:color,weight:700}));
  });box.appendChild(s);
}
function renderInspect(box){
  box.innerHTML="";const W=box.clientWidth,H=box.clientHeight;if(!W||!H)return;
  const p={t:30,r:44,b:26,l:40,h:H-56};const s=svgEl(W,H);
  const lbs=DATA.months,n=lbs.length,iw=W-p.l-p.r;
  const maxV=niceMax(Math.max(...DATA.totals.in_inspection),3);
  const allRates=[...DATA.inspection_rate.in,...DATA.inspection_rate.out].filter(x=>!isNaN(x)&&x>0);
  const maxR=Math.max(1,Math.ceil(Math.max(...allRates)));
  gridY(s,p,p.l,W-p.r,maxV,4,v=>fmt(v));const band=iw/n,bw=Math.min(11,band*0.3);
  const X=i=>p.l+band*(i+.5),Y=v=>p.t+p.h*(1-v/maxV),YR=r=>p.t+p.h*(1-r/maxR);
  lbs.forEach((l,i)=>s.appendChild(text(X(i),H-8,l,{anchor:"middle",size:9.5})));
  DATA.totals.in_inspection.forEach((v,i)=>{const x=X(i)-bw-1.5,h=p.h*v/maxV;const r=mk("rect",{x,y:p.t+p.h-h,width:bw,height:h,fill:C.violet,rx:2,opacity:.92});r.appendChild(mk("title",{})).textContent=`${lbs[i]} 入仓查验 ${v} 次`;s.appendChild(r)});
  DATA.totals.out_inspection.forEach((v,i)=>{const x=X(i)+1.5,h=p.h*v/maxV;const r=mk("rect",{x,y:p.t+p.h-h,width:bw,height:h,fill:"#4f7fa8",rx:2,opacity:.9});r.appendChild(mk("title",{})).textContent=`${lbs[i]} 出仓查验 ${v} 次`;s.appendChild(r)});
  // 入仓查验率（mint 实线）
  const pIn=DATA.inspection_rate.in.map((v,i)=>[X(i),YR(v)]);
  s.appendChild(mk("path",{d:smoothPath(pIn),fill:"none",stroke:C.mint,"stroke-width":2}));
  pIn.forEach((pt,i)=>{const c=mk("circle",{cx:pt[0],cy:pt[1],r:2.8,fill:"#04121f",stroke:C.mint,"stroke-width":1.8});c.appendChild(mk("title",{})).textContent=`${lbs[i]} 入仓查验率 ${DATA.inspection_rate.in[i]}%`;s.appendChild(c)});
  // 出仓查验率（cyan 虚线）
  const pOut=DATA.inspection_rate.out.map((v,i)=>[X(i),YR(v)]);
  s.appendChild(mk("path",{d:smoothPath(pOut),fill:"none",stroke:C.cyan,"stroke-width":2,"stroke-dasharray":"4 3"}));
  pOut.forEach((pt,i)=>{const c=mk("circle",{cx:pt[0],cy:pt[1],r:2.8,fill:"#04121f",stroke:C.cyan,"stroke-width":1.8});c.appendChild(mk("title",{})).textContent=`${lbs[i]} 出仓查验率 ${DATA.inspection_rate.out[i]}%`;s.appendChild(c)});
  // 右侧 % 刻度
  [0,.5,1].forEach(t=>s.appendChild(text(W-p.r+8,YR(maxR*t)+3.5,(maxR*t).toFixed(1)+"%",{size:9.5,fill:C.mint})));
  const lg=[["入仓查验",C.violet],["出仓查验","#4f7fa8"],["入仓查验率",C.mint,false],["出仓查验率",C.cyan,true]];let lx=p.l;
  lg.forEach(([n2,c,dash])=>{
    s.appendChild(mk(dash?"line":"rect",dash?{x1:lx,y1:14,x2:lx+8,y2:14,stroke:c,"stroke-width":2,"stroke-dasharray":"2 2"}:{x:lx,y:10,width:8,height:8,rx:2,fill:c}));
    s.appendChild(text(lx+12,17.5,n2,{size:10,fill:"#9dc0d8"}));lx+=n2.length*10+34
  });box.appendChild(s);
}
function renderDiff(box){
  box.innerHTML="";const W=box.clientWidth,H=box.clientHeight;if(!W||!H)return;
  const p={t:24,r:16,b:26,l:46,h:H-50};const s=svgEl(W,H);
  const lbs=DATA.months,n=lbs.length,iw=W-p.l-p.r;
  const max=niceMax(Math.max(...DATA.totals.diff),4);gridY(s,p,p.l,W-p.r,max,4,fmt);
  const band=iw/n,bw=Math.min(30,band*0.46);const showVal=band>=46;
  DATA.totals.diff.forEach((v,i)=>{const x=p.l+band*i+(band-bw)/2,h=p.h*v/max;const isLast=i===n-1;
    const grad=mk("linearGradient",{id:"dg"+i,x1:0,y1:0,x2:0,y2:1});
    grad.appendChild(mk("stop",{offset:"0%","stop-color":isLast?"#ff8a97":C.red,"stop-opacity":.95}));
    grad.appendChild(mk("stop",{offset:"100%","stop-color":C.red,"stop-opacity":.28}));
    const defs=mk("defs",{});defs.appendChild(grad);s.appendChild(defs);
    const r=mk("rect",{x,y:p.t+p.h-h,width:bw,height:h,fill:`url(#dg${i})`,rx:3});
    r.appendChild(mk("title",{})).textContent=`${lbs[i]} 库存差异 ${fmt(v)}`;s.appendChild(r);
    if(showVal||isLast){const ty=Math.max(11,p.t+p.h-h-5);s.appendChild(text(x+bw/2,ty,fmt(v),{anchor:"middle",size:9.5,fill:isLast?"#ffb6c0":"#b9cddd"}))}
    s.appendChild(text(x+bw/2,H-8,lbs[i],{anchor:"middle",size:9.5}))});box.appendChild(s);
}
function renderCoo(box){
  box.innerHTML="";const W=box.clientWidth,H=box.clientHeight;if(!W||!H)return;
  const s=svgEl(W,H);const items=DATA.cooOct,total=items.reduce((a,b)=>a+b[1],0);
  const cx=W/2,cy=H/2,R=Math.min(H*0.44,W*0.44),r=R*0.62;let ang=-Math.PI/2;
  items.forEach(([name,v],i)=>{const a2=ang+Math.PI*2*v/total;
    const p1=[cx+R*Math.cos(ang),cy+R*Math.sin(ang)],p2=[cx+R*Math.cos(a2),cy+R*Math.sin(a2)];
    const q1=[cx+r*Math.cos(a2),cy+r*Math.sin(a2)],q2=[cx+r*Math.cos(ang),cy+r*Math.sin(ang)];
    const large=(a2-ang)>Math.PI?1:0;
    const d=`M${p1[0].toFixed(2)},${p1[1].toFixed(2)} A${R},${R} 0 ${large} 1 ${p2[0].toFixed(2)},${p2[1].toFixed(2)} L${q1[0].toFixed(2)},${q1[1].toFixed(2)} A${r},${r} 0 ${large} 0 ${q2[0].toFixed(2)},${q2[1].toFixed(2)} Z`;
    const path=mk("path",{d,fill:COO_COLORS[i%COO_COLORS.length],opacity:i===0?1:.86,stroke:"#08243a","stroke-width":1.5});
    path.appendChild(mk("title",{})).textContent=`${name}：${v} 份（${(v/total*100).toFixed(1)}%）`;
    s.appendChild(path);ang=a2});
  s.appendChild(text(cx,cy+2,total+"",{anchor:"middle",size:Math.max(18,R*0.5),fill:"#fff",weight:700}));
  s.appendChild(text(cx,cy+R*0.34,"10月认证份数",{anchor:"middle",size:Math.max(9,R*0.16),fill:"#8fb2ca"}));
  box.appendChild(s);
  const ytd=new Map(DATA.cooYtd);
  document.getElementById("coo-legend").innerHTML=items.map(([name,v],i)=>`
    <div class="row"><span class="dot" style="background:${COO_COLORS[i%COO_COLORS.length]}"></span>
    <span class="nm" title="累计 ${ytd.get(name)||0} 份">${name}</span>
    <span class="vl">${v}</span><span class="pc">${(v/total*100).toFixed(0)}%</span></div>`).join("");
}
function renderAmount(box){
  box.innerHTML="";const W=box.clientWidth,H=box.clientHeight;if(!W||!H)return;
  const s=svgEl(W,H);
  const inA=DATA.totals.amount_in||[],outA=DATA.totals.amount_out||[];
  const hasData=inA.some(v=>v>0)||outA.some(v=>v>0);
  if(!hasData){box.appendChild(text(W/2,H/2,"暂无进出口金额数据",{anchor:"middle",size:12,fill:"#7fa3bf"}));return}
  const p={t:26,r:16,b:26,l:60,h:H-52};
  const lbs=DATA.months,n=lbs.length,iw=W-p.l-p.r;
  const max=niceMax(Math.max(...inA.filter(x=>x>0),...outA.filter(x=>x>0)),4);
  gridY(s,p,p.l,W-p.r,max,4,v=>fmt(v));
  const band=iw/n,bw=Math.min(13,band*0.32);const X=i=>p.l+band*(i+.5),Y=v=>p.t+p.h*(1-v/max);
  lbs.forEach((l,i)=>s.appendChild(text(X(i),H-8,l,{anchor:"middle",size:9.5})));
  inA.forEach((v,i)=>{if(v<=0)return;const x=X(i)-bw-1,h=p.h*v/max;
    const grad=mk("linearGradient",{id:"ai"+i,x1:0,y1:0,x2:0,y2:1});
    grad.appendChild(mk("stop",{offset:"0%","stop-color":C.cyan,"stop-opacity":.95}));
    grad.appendChild(mk("stop",{offset:"100%","stop-color":C.cyan,"stop-opacity":.35}));
    const d=mk("defs",{});d.appendChild(grad);s.appendChild(d);
    const r=mk("rect",{x,y:p.t+p.h-h,width:bw,height:h,fill:`url(#ai${i})`,rx:2});
    r.appendChild(mk("title",{})).textContent=`${lbs[i]} 入仓金额 ${fmt(v)}`;s.appendChild(r)});
  outA.forEach((v,i)=>{if(v<=0)return;const x=X(i)+1,h=p.h*v/max;
    const grad=mk("linearGradient",{id:"ao"+i,x1:0,y1:0,x2:0,y2:1});
    grad.appendChild(mk("stop",{offset:"0%","stop-color":C.amber,"stop-opacity":.95}));
    grad.appendChild(mk("stop",{offset:"100%","stop-color":C.amber,"stop-opacity":.35}));
    const d=mk("defs",{});d.appendChild(grad);s.appendChild(d);
    const r=mk("rect",{x,y:p.t+p.h-h,width:bw,height:h,fill:`url(#ao${i})`,rx:2});
    r.appendChild(mk("title",{})).textContent=`${lbs[i]} 出仓金额 ${fmt(v)}`;s.appendChild(r)});
  const lg=[["入仓金额",C.cyan],["出仓金额",C.amber]];let lx=p.l;
  lg.forEach(([n2,c])=>{s.appendChild(mk("rect",{x:lx,y:10,width:8,height:8,rx:2,fill:c}));s.appendChild(text(lx+12,17.5,n2,{size:10,fill:"#9dc0d8"}));lx+=n2.length*10+46});
  // 主视角月份标注
  const lastIn=inA[n-1]||0,lastOut=outA[n-1]||0;
  const tip=text(X(n-1),p.t-4,`本月入 ${fmt(lastIn)}  出 ${fmt(lastOut)}`,{anchor:"middle",size:10,fill:"#cfe8f7",weight:600});s.appendChild(tip);
  box.appendChild(s);
}
function renderLoc(box){
  box.innerHTML="";const W=box.clientWidth,H=box.clientHeight;if(!W||!H)return;
  const s=svgEl(W,H);const items=DATA.locs,total=items.reduce((a,b)=>a+b[1],0);
  const p={t:22,l:76,r:40,b:6};const max=Math.max(...items.map(x=>x[1]));
  const rowH=(H-p.t-p.b)/items.length;
  items.forEach(([name,v],i)=>{const y=p.t+rowH*i+rowH/2;const bw=(W-p.l-p.r)*v/max;
    s.appendChild(text(p.l-10,y+4,name,{anchor:"end",size:10.5,fill:"#b9cddd"}));
    const grad=mk("linearGradient",{id:"lg"+i,x1:0,y1:0,x2:1,y2:0});
    grad.appendChild(mk("stop",{offset:"0%","stop-color":"#1f6f95"}));
    grad.appendChild(mk("stop",{offset:"100%","stop-color":"#6fd0ff"}));
    const defs=mk("defs",{});defs.appendChild(grad);s.appendChild(defs);
    const r=mk("rect",{x:p.l,y:y-6,width:Math.max(bw,2),height:12,fill:`url(#lg${i})`,rx:4});
    r.appendChild(mk("title",{})).textContent=`${name}：${v} 次`;s.appendChild(r);
    const label=`${v} 次  ${(v/total*100).toFixed(0)}%`;const tx=p.l+bw+8;
    if(tx+label.length*9+4>W-4)s.appendChild(text(p.l+bw-7,y+4,label,{anchor:"end",size:10,fill:"#04121f",weight:600}));
    else s.appendChild(text(tx,y+4,label,{size:10,fill:"#9dc0d8"}))});box.appendChild(s);
}
function renderKpi(){
  const defs=[
    {label:"入仓报关单份数",k:"inbound",color:"var(--cyan)",unit:"份"},
    {label:"出仓报关单份数",k:"outbound",color:"var(--amber)",unit:"份"},
    {label:"入仓查验次数",k:"in_inspection",color:"var(--violet)",unit:"次",rate:true},
    {label:"出仓查验次数",k:"out_inspection",color:"#6f8ff0",unit:"次",rate:true},
    {label:"COO证书认证",k:"coo",color:"var(--mint)",unit:"份"},
    {label:"库存差异",k:"diff",color:"var(--red)",unit:"",rate:false}
  ];
  const host=document.getElementById("kpis");
  host.innerHTML=defs.map(d=>{
    const v=DATA.kpi[d.k];const up=v.mom>=0;const arrow=up?"▲":"▼";const cls=up?"up":"down";
    const foot=d.rate?`查验率 <b style="color:#dceefb">${v.rate}%</b>`:`累计 <b style="color:#dceefb">${fmt(v.ytd)}</b>${d.unit}`;
    return `<div class="kpi" style="--accent:${d.color}">
      <div class="k-label"><span>${d.label}</span><span class="k-tag">10月</span></div>
      <div class="k-val">${fmt(v.value)}<small>${d.unit}</small></div>
      <div class="k-foot"><span class="delta ${cls}">${arrow} ${Math.abs(v.mom).toFixed(1)}%</span><span>环比</span><span class="sep"></span><span>${foot}</span></div>
    </div>`}).join("");
}
const charts=[["ch-trend",renderTrend],["ch-inspect",renderInspect],["ch-diff",renderDiff],["ch-coo",renderCoo],["ch-amount",renderAmount],["ch-loc",renderLoc]];
function renderAll(){renderKpi();charts.forEach(([id,fn])=>{const box=document.getElementById(id);if(box)fn(box)})}
let rt;window.addEventListener("resize",()=>{clearTimeout(rt);rt=setTimeout(renderAll,160)});renderAll();
function tick(){const d=new Date();const pad=x=>String(x).padStart(2,"0");document.getElementById("clock").textContent=`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`}
tick();setInterval(tick,1000);
</script>
</body>
</html>
'''

def render_html(data, src_path, out_path):
    main_month_zh = str(MAIN_MONTH_IDX + 1) + "月" if MAIN_MONTH_IDX < 9 else \
                   ("10月" if MAIN_MONTH_IDX == 9 else
                    ("11月" if MAIN_MONTH_IDX == 10 else "12月"))
    month1_zh = "1月"
    period_zh = f"2026年 {month1_zh}-{main_month_zh}"

    # 额外准备 HTML 里要用的几个派生字段
    coo_oct = data["top"]["coo"][:5]
    coo_oct_names = [x["name"] for x in coo_oct]
    coo_oct_rest = sum(x["value"] for x in data["top"]["coo"] if x["name"] not in coo_oct_names)
    coo_oct_final = coo_oct + ([{"name": "其他", "value": coo_oct_rest}] if coo_oct_rest else [])

    payload = {
        "period": period_zh,
        "months":  [f"{i+1}月" for i in range(len(data["months"]))],
        "totals":  data["totals"],
        "kpi":     data["kpi"],
        "top": {
            "inbound":  [[x["name"], int(x["value"])] for x in data["top"]["inbound"]],
            "outbound": [[x["name"], int(x["value"])] for x in data["top"]["outbound"]],
        },
        "cooOct": [[x["name"], int(x["value"])] for x in coo_oct_final],
        "cooYtd": [[x["name"], int(x["value"])] for x in data["coo_ytd"]],
        "inspection_rate": data["inspection_rate"],
        "locs": [[n, int(v)] for n, v in data["in_inspection_locs"]],
    }

    # footer 里金额来源标注
    amt_src = data.get("amount_source")
    if amt_src:
        if os.path.basename(amt_src) == os.path.basename(src_path):
            amt_label = "、进出口金额"
        else:
            amt_label = f"，进出口金额来自 {os.path.basename(amt_src)}"
    else:
        amt_label = ""

    src_basename = os.path.basename(src_path)
    html = (HTML_TEMPLATE
            .replace("{{PERIOD}}",    period_zh)
            .replace("{{MAIN_LABEL}}",f"2026年{main_month_zh}")
            .replace("{{SRC_BASENAME}}", src_basename)
            .replace("{{AMOUNT_UNIT}}",  AMOUNT_UNIT)
            .replace("{{AMOUNT_SRC_LABEL}}", amt_label)
            .replace("{{DATA_JSON}}", json.dumps(payload, ensure_ascii=False)))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)


# ---------- main ----------
def main():
    print("=" * 60)
    print("关务月报数据大屏 · 一键生成")
    print("=" * 60)
    if not os.path.exists(SRC):
        print(f"[错误] 找不到源文件: {os.path.abspath(SRC)}")
        print("       请把脚本放在月报同级目录，或修改脚本顶部的 SRC 配置")
        sys.exit(1)

    # 输出目录：优先放在源文件同目录；若不可写则落在脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.dirname(os.path.abspath(SRC))
    out_dir = src_dir if os.access(src_dir, os.W_OK) else script_dir
    if out_dir != src_dir:
        print(f"      （源文件目录不可写，输出改到脚本目录: {out_dir}）")

    base_no_ext = os.path.splitext(os.path.basename(SRC))[0]
    out_xlsx = os.path.join(out_dir, f"{base_no_ext}_{OUT_XLSX_SUFFIX}")
    out_html = os.path.join(out_dir, f"{base_no_ext}_{OUT_HTML_SUFFIX}")

    # 1 取数
    print(f"\n[1/3] 从 {os.path.basename(SRC)} 取数 ...")
    data = extract(SRC, MAIN_MONTH_IDX)
    print(f"      主视角 KPI（{data['main_month']}）:")
    for k, v in data["kpi"].items():
        if k.startswith("amount_"): continue   # 金额 KPI 单独打印
        arrow = "▲" if (v.get("mom") or 0) >= 0 else "▼"
        print(f"        {k:14s} {v['value']:>8}  {arrow} {abs(v.get('mom') or 0):.1f}%  累计 {v['ytd']}")
    # 进出口金额
    amt_src = data.get("amount_source")
    if amt_src:
        a_in  = data["kpi"]["amount_in"]["value"]
        a_out = data["kpi"]["amount_out"]["value"]
        print(f"        {'amount_in':14s} {a_in:>12,}  (来源: {amt_src})")
        print(f"        {'amount_out':14s} {a_out:>12,}")
    else:
        print(f"        ⚠ 未找到进出口金额数据（在 SRC 或 AMOUNT_SRC 里没读到）。"
              "可检查 CONFIG 的 AMOUNT_SRC 路径，或把金额表放入 SRC 并命名为“进出口金额”。")

    # 2 Excel 原生看板
    print(f"\n[2/3] 生成 Excel 原生看板 -> {os.path.basename(out_xlsx)}")
    try:
        build_excel(data, SRC, out_xlsx)
        print("      ✓ 已写入「数据大屏」子表，其他子表全部保留")
    except ModuleNotFoundError:
        print("      ✗ 需要先安装: pip install pywin32")
        raise
    except Exception as e:
        print(f"      ✗ Excel 看板生成失败: {e}")
        traceback.print_exc()

    # 3 HTML 大屏
    print(f"\n[3/3] 生成 HTML 全屏大屏 -> {os.path.basename(out_html)}")
    try:
        render_html(data, SRC, out_html)
        print("      ✓ 深色科技风，浏览器打开即可全屏演示")
    except Exception as e:
        print(f"      ✗ HTML 生成失败: {e}")
        traceback.print_exc()

    print("\n" + "=" * 60)
    print("完成！")
    print(f"  Excel 看板: {out_xlsx}")
    print(f"  HTML 大屏:  {out_html}")
    print("=" * 60)


if __name__ == "__main__":
    main()
