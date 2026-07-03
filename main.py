"""
INT / POS Algo Calculator  (Flask web app)
==========================================

A small web app that:
  1. Takes a data CSV (same columns as data.csv).
  2. Takes a "date from" and "date to" range (optional -> defaults to full range).
  3. Classifies every account (alias, per algo) as INT or POS based on DTE behaviour:
        - only 0DTE across the whole range           -> INT (intraday)
        - any non-zero DTE (1DTE, 4DTE, 1DTE_3DTE ...) -> POS (positional)
  4. Shows a Summary + per-algo/type tables on the page, and offers an Excel
     download with one sheet per "Algo<n> INT" / "Algo<n> POS" plus a Summary sheet.
  5. Filtering system with search boxes for User ID, Broker, Algo, and Server.
     Also includes "Only show portfolios containing 'QS'" checkbox.
"""

import html as html_lib
import io
import re
import uuid
from itertools import groupby

import pandas as pd
from flask import Flask, request, render_template_string, send_file, abort, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

DTE_NUM_RE = re.compile(r"(\d+)\s*dte", re.IGNORECASE)


def dte_numbers(value) -> list[int]:
    """Extract all numeric DTE values from a dte cell."""
    if pd.isna(value):
        return []
    return [int(n) for n in DTE_NUM_RE.findall(str(value))]


def normalize_target_user_id(value):
    """Collapse known leading-zero variants of a user_id onto one canonical form."""
    normalized = str(value).strip()

    if normalized in {"06954037", "6954037"}:
        return "06954037"

    if normalized in {"04101961", "4101961"}:
        return "04101961"

    return normalized


def classify_type(dte_series: pd.Series) -> str:
    """INT if every row is 0DTE only; POS if any non-zero DTE appears."""
    for v in dte_series:
        nums = dte_numbers(v)
        if nums and any(n != 0 for n in nums):
            return "POS"
    return "INT"


def load_csv(file_bytes: bytes) -> pd.DataFrame:
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding="latin-1")
    df.columns = [c.replace("﻿", "").strip() for c in df.columns]
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    for col in ("mtm_all", "allocation", "cumulative_mtm"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "user_id" in df.columns:
        df["user_id"] = df["user_id"].map(
            lambda v: v if pd.isna(v) else normalize_target_user_id(v))
    return df


def build_account_rollup(df: pd.DataFrame) -> pd.DataFrame:
    """One row per (algo, user_id): classification + aggregated numbers."""
    rows = []
    df_sorted = df.sort_values("date")
    for (algo, uid), g in df_sorted.groupby(["algo", "user_id"], sort=False):
        acct_type = classify_type(g["dte"])
        last = g.iloc[-1]
        aum = g["allocation"].mean() * 100
        total_mtm = g["mtm_all"].sum()
        mtm_pct = (total_mtm / aum * 100) if aum else float("nan")
        rows.append(
            {
                "Algo": algo,
                "Type": acct_type,
                "Alias": last.get("alias"),
                "UserID": uid,
                "AUM": aum,
                "Total MTM": total_mtm,
                "MTM%": mtm_pct,
                "Trading Days": len(g),
                "Broker": last.get("broker"),
                "Server": last.get("server"),
            }
        )
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["Algo", "Type", "Total MTM"], ascending=[True, True, False])
    return out


def build_summary(rollup: pd.DataFrame) -> pd.DataFrame:
    """One row per (Algo, Type) with totals + a grand-total row."""
    if rollup.empty:
        return pd.DataFrame()
    grp = rollup.groupby(["Algo", "Type"], as_index=False).agg(
        Accounts=("Alias", "count"),
        Total_AUM=("AUM", "sum"),
        Total_MTM=("Total MTM", "sum"),
    )
    grp["MTM%"] = grp.apply(
        lambda r: (r["Total_MTM"] / r["Total_AUM"] * 100) if r["Total_AUM"] else float("nan"),
        axis=1,
    )
    grp = grp.rename(columns={"Total_AUM": "Total AUM", "Total_MTM": "Total MTM", "Accounts": "Accounts"})
    grp = grp[["Algo", "Type", "Accounts", "Total AUM", "Total MTM", "MTM%"]]
    grp = grp.sort_values(["Algo", "Type"]).reset_index(drop=True)

    total_aum = grp["Total AUM"].sum()
    total_mtm = grp["Total MTM"].sum()
    grand = {
        "Algo": "TOTAL",
        "Type": "",
        "Accounts": grp["Accounts"].sum(),
        "Total AUM": total_aum,
        "Total MTM": total_mtm,
        "MTM%": (total_mtm / total_aum * 100) if total_aum else float("nan"),
    }
    return pd.concat([grp, pd.DataFrame([grand])], ignore_index=True)


def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name))
    return name[:31]


# --- Excel styling / formula helpers ---------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F2A44")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)
TOTAL_TOP = Border(top=Side(style="thin", color="38BDF8"))
RED_FONT = Font(color="D12F2F")
GREEN_FONT = Font(color="1E9E5A")
CENTER = Alignment(horizontal="center")

MONEY_FMT = "#,##0"
PCT_FMT = '0.00"%"'

ALGO_BANDS_XL = [
    ("F3D9D1", "E6B3A5"),
    ("DCEBD6", "BCD8B3"),
    ("E5DDF3", "C8BCE8"),
    ("D7E7F4", "B0D0EC"),
    ("D5EDE5", "AAD9C9"),
    ("F6E6CF", "ECD0A4"),
]
GRAND_FILL = PatternFill("solid", fgColor="D9DEE8")
ALGO_BANDS_WEB = [
    ("#f5ddd5", "#ecc2b5"),
    ("#deecd8", "#c3dcbb"),
    ("#e7e0f4", "#d0c4ec"),
    ("#dae9f5", "#b9d6ef"),
    ("#d9efe8", "#b6ddce"),
    ("#f7ead2", "#eed6ab"),
]


def _cell_val(v):
    """Clean a pandas value for openpyxl."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(v, "item"):
        try:
            return v.item()
        except (ValueError, TypeError):
            return v
    return v


def _algo_criteria(v):
    """Excel criteria literal for an algo value."""
    try:
        fv = float(v)
        return str(int(fv)) if fv.is_integer() else repr(fv)
    except (TypeError, ValueError):
        return '"' + str(v).replace('"', '""') + '"'


def _style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
    ws.freeze_panes = "A2"


def _add_sign_colors(ws, ranges):
    """Red font for negative values, green for positive."""
    for rng in ranges:
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="lessThan", formula=["0"], font=RED_FONT))
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN_FONT))


def build_excel(rollup: pd.DataFrame, raw: pd.DataFrame) -> bytes:
    """Build the workbook with every figure as a live Excel formula."""
    buf = io.BytesIO()
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"

    ws_data = wb.create_sheet("Data")
    ws_data.append(["Date", "Algo", "Alias", "UserID", "Allocation", "MTM_all"])
    sub = raw.reindex(columns=["date", "algo", "alias", "user_id", "allocation", "mtm_all"])
    for row in sub.itertuples(index=False):
        d = row.date
        d = d.to_pydatetime() if pd.notna(d) else None
        ws_data.append([d, _cell_val(row.algo), _cell_val(row.alias), _cell_val(row.user_id),
                        _cell_val(row.allocation), _cell_val(row.mtm_all)])
    data_last = len(sub) + 1
    for r in range(2, data_last + 1):
        ws_data.cell(r, 1).number_format = "yyyy-mm-dd"
        ws_data.cell(r, 5).number_format = MONEY_FMT
        ws_data.cell(r, 6).number_format = MONEY_FMT
    _style_header(ws_data, 6)
    for col, w in zip("ABCDEF", (13, 10, 22, 12, 16, 16)):
        ws_data.column_dimensions[col].width = w

    algo_rng = f"Data!$B$2:$B${data_last}"
    uid_rng = f"Data!$D$2:$D${data_last}"
    alloc_rng = f"Data!$E$2:$E${data_last}"
    mtm_rng = f"Data!$F$2:$F${data_last}"

    headers = ["Alias", "UserID", "AUM", "Total MTM", "MTM%",
               "Trading Days", "Broker", "Server"]
    summary_rows = []

    if not rollup.empty:
        for (algo, acct_type), g in rollup.groupby(["Algo", "Type"], sort=True):
            sheet = safe_sheet_name(f"Algo{algo} {acct_type}")
            ws = wb.create_sheet(sheet)
            ws.append(headers)
            algo_lit = _algo_criteria(algo)

            r = 2
            for _, rec in g.iterrows():
                ws.cell(r, 1, _cell_val(rec["Alias"]))
                ws.cell(r, 2, _cell_val(rec.get("UserID")))
                ws.cell(r, 3, (f"=IFERROR(AVERAGEIFS({alloc_rng},{algo_rng},"
                               f"{algo_lit},{uid_rng},$B{r})*100,0)"))
                ws.cell(r, 4, (f"=SUMIFS({mtm_rng},{algo_rng},{algo_lit},"
                               f"{uid_rng},$B{r})"))
                ws.cell(r, 5, f'=IF(C{r}=0,"",D{r}/C{r}*100)')
                ws.cell(r, 6, _cell_val(rec.get("Trading Days")))
                ws.cell(r, 7, _cell_val(rec.get("Broker")))
                ws.cell(r, 8, _cell_val(rec.get("Server")))
                r += 1

            tr = r
            ws.cell(tr, 1, "TOTAL")
            ws.cell(tr, 3, f"=SUM(C2:C{tr - 1})")
            ws.cell(tr, 4, f"=SUM(D2:D{tr - 1})")
            ws.cell(tr, 5, f'=IF(C{tr}=0,"",D{tr}/C{tr}*100)')

            for rr in range(2, tr + 1):
                ws.cell(rr, 3).number_format = MONEY_FMT
                ws.cell(rr, 4).number_format = MONEY_FMT
                ws.cell(rr, 5).number_format = PCT_FMT
            for c in range(1, 9):
                ws.cell(tr, c).font = TOTAL_FONT
                ws.cell(tr, c).border = TOTAL_TOP
            _style_header(ws, 8)
            _add_sign_colors(ws, [f"D2:D{tr}", f"E2:E{tr}"])
            for col, w in zip("ABCDEFGH", (20, 16, 14, 14, 11, 13, 16, 14)):
                ws.column_dimensions[col].width = w

            summary_rows.append((algo, acct_type, len(g), sheet, tr))

    ws_sum.append(["Algo", "Type", "Accounts", "Total AUM", "Total MTM", "MTM%"])
    r = 2
    for band_i, (algo, items_iter) in enumerate(
            groupby(summary_rows, key=lambda x: x[0])):
        items = list(items_iter)
        light, mid = ALGO_BANDS_XL[band_i % len(ALGO_BANDS_XL)]
        light_fill = PatternFill("solid", fgColor=light)
        mid_fill = PatternFill("solid", fgColor=mid)
        block_start = r
        for _algo, acct_type, count, sheet, tr in items:
            ws_sum.cell(r, 1, _cell_val(algo))
            ws_sum.cell(r, 2, acct_type)
            ws_sum.cell(r, 3, count)
            ws_sum.cell(r, 4, f"='{sheet}'!C{tr}")
            ws_sum.cell(r, 5, f"='{sheet}'!D{tr}")
            ws_sum.cell(r, 6, f'=IF(D{r}=0,"",E{r}/D{r}*100)')
            for c in range(1, 7):
                ws_sum.cell(r, c).fill = light_fill
            r += 1
        block_end = r - 1
        ws_sum.cell(r, 1, _cell_val(algo))
        ws_sum.cell(r, 2, "Total")
        ws_sum.cell(r, 3, f"=SUM(C{block_start}:C{block_end})")
        ws_sum.cell(r, 4, f"=SUM(D{block_start}:D{block_end})")
        ws_sum.cell(r, 5, f"=SUM(E{block_start}:E{block_end})")
        ws_sum.cell(r, 6, f'=IF(D{r}=0,"",E{r}/D{r}*100)')
        for c in range(1, 7):
            ws_sum.cell(r, c).fill = mid_fill
            ws_sum.cell(r, c).font = TOTAL_FONT
        r += 1

    if summary_rows:
        data_last = r - 1
        b_rng = f"$B$2:$B${data_last}"
        c_rng = f"$C$2:$C${data_last}"
        d_rng = f"$D$2:$D${data_last}"
        e_rng = f"$E$2:$E${data_last}"
        int_row = r
        ws_sum.cell(r, 1, "TOTAL INT")
        ws_sum.cell(r, 3, f'=SUMIF({b_rng},"INT",{c_rng})')
        ws_sum.cell(r, 4, f'=SUMIF({b_rng},"INT",{d_rng})')
        ws_sum.cell(r, 5, f'=SUMIF({b_rng},"INT",{e_rng})')
        ws_sum.cell(r, 6, f'=IF(D{r}=0,"",E{r}/D{r}*100)')
        r += 1
        pos_row = r
        ws_sum.cell(r, 1, "TOTAL POS")
        ws_sum.cell(r, 3, f'=SUMIF({b_rng},"POS",{c_rng})')
        ws_sum.cell(r, 4, f'=SUMIF({b_rng},"POS",{d_rng})')
        ws_sum.cell(r, 5, f'=SUMIF({b_rng},"POS",{e_rng})')
        ws_sum.cell(r, 6, f'=IF(D{r}=0,"",E{r}/D{r}*100)')
        r += 1
        ws_sum.cell(r, 1, "TOTAL OVERALL")
        ws_sum.cell(r, 3, f"=C{int_row}+C{pos_row}")
        ws_sum.cell(r, 4, f"=D{int_row}+D{pos_row}")
        ws_sum.cell(r, 5, f"=E{int_row}+E{pos_row}")
        ws_sum.cell(r, 6, f'=IF(D{r}=0,"",E{r}/D{r}*100)')
        grand_last = r

        for rr in range(int_row, grand_last + 1):
            for c in range(1, 7):
                cell = ws_sum.cell(rr, c)
                cell.fill = GRAND_FILL
                cell.font = TOTAL_FONT
                cell.border = TOTAL_TOP
        for rr in range(2, grand_last + 1):
            ws_sum.cell(rr, 3).number_format = "#,##0"
            ws_sum.cell(rr, 4).number_format = MONEY_FMT
            ws_sum.cell(rr, 5).number_format = MONEY_FMT
            ws_sum.cell(rr, 6).number_format = PCT_FMT
        _add_sign_colors(ws_sum, [f"E2:E{grand_last}", f"F2:F{grand_last}"])
    _style_header(ws_sum, 6)
    for col, w in zip("ABCDEF", (14, 8, 12, 18, 18, 10)):
        ws_sum.column_dimensions[col].width = w

    wb.move_sheet(ws_data, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("Data"))
    wb.active = 0

    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

_EXCEL_CACHE: dict[str, tuple[str, bytes]] = {}
_DATA_CACHE: dict[str, tuple[str, bytes]] = {}
_FILTER_CACHE: dict[str, tuple[pd.DataFrame, dict]] = {}  # token -> (df, filter_state)


def _skey(x) -> str:
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _unique_options(df: pd.DataFrame, col: str) -> list[str]:
    if col not in df.columns:
        return []
    vals = (_skey(v) for v in df[col].dropna())
    uniq = [v for v in dict.fromkeys(vals) if v and v.lower() != "nan"]

    def key(x):
        try:
            return (0, float(x))
        except (TypeError, ValueError):
            return (1, x.lower())

    return sorted(uniq, key=key)


def _filter_items_html(options, selected, name: str, filter_type: str = "") -> str:
    """Checkbox items for a filter panel."""
    parts = []
    for v in options:
        chk = "checked" if (selected is None or v in selected) else ""
        safe = html_lib.escape(str(v))
        # Add data attribute for filtering
        parts.append(
            f'<label class="fitem" data-filter-type="{filter_type}">'
            f'<input type="checkbox" name="{name}" value="{safe}" {chk}> '
            f'<span>{safe}</span></label>'
        )
    return "".join(parts)


PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>INT / POS Algo Calculator</title>
  <style>
    :root {
      --bg:#0b1120; --bg2:#0f172a; --card:#151f34; --card2:#1a2540;
      --ink:#e8eefc; --muted:#8ea0bf; --accent:#38bdf8; --accent2:#818cf8;
      --green:#34d399; --red:#fb7185; --line:#25324e; --line2:#334155;
    }
    * { box-sizing:border-box; }
    html { scroll-behavior:smooth; }
    body {
      margin:0; font-family:'Inter',system-ui,Segoe UI,Roboto,Arial,sans-serif;
      color:var(--ink); min-height:100vh;
      background:
        radial-gradient(900px 500px at 12% -8%, rgba(56,189,248,.10), transparent 60%),
        radial-gradient(760px 460px at 100% 0%, rgba(129,140,248,.12), transparent 55%),
        var(--bg);
    }
    .wrap { max-width:1240px; margin:0 auto; padding:36px 22px 72px; }

    .head { display:flex; align-items:center; gap:14px; margin-bottom:6px; }
    .logo {
      width:46px; height:46px; border-radius:13px; display:grid; place-items:center;
      font-size:24px; background:linear-gradient(135deg,var(--accent),var(--accent2));
      box-shadow:0 8px 22px rgba(56,189,248,.28);
    }
    h1 {
      font-size:26px; margin:0; letter-spacing:-.4px; font-weight:800;
      background:linear-gradient(90deg,#fff,#a9c7ff);
      -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent;
    }
    .sub { color:var(--muted); margin:0 0 26px 60px; font-size:14px; }

    .card {
      background:linear-gradient(180deg,var(--card),var(--bg2));
      border:1px solid var(--line); border-radius:16px;
      padding:22px; margin-bottom:24px;
      box-shadow:0 1px 0 rgba(255,255,255,.03) inset, 0 18px 40px -24px rgba(0,0,0,.7);
    }

    form { display:flex; flex-wrap:wrap; gap:18px; align-items:end; }
    .field { display:flex; flex-direction:column; }
    label { font-size:12px; color:var(--muted); margin-bottom:7px; font-weight:600;
            text-transform:uppercase; letter-spacing:.5px; }
    input[type=file], input[type=date] {
      background:#0a1122; border:1px solid var(--line2); color:var(--ink);
      border-radius:10px; padding:11px 12px; font-size:14px; transition:border-color .15s, box-shadow .15s;
    }
    input[type=file]:hover, input[type=date]:hover { border-color:#44557a; }
    input[type=file]:focus, input[type=date]:focus {
      outline:0; border-color:var(--accent); box-shadow:0 0 0 3px rgba(56,189,248,.18);
    }
    input[type=file]::file-selector-button {
      background:#1c2b48; color:var(--ink); border:0; border-radius:7px;
      padding:7px 12px; margin-right:12px; cursor:pointer; font-weight:600;
    }
    button {
      background:linear-gradient(135deg,var(--accent),var(--accent2)); color:#051426;
      border:0; border-radius:10px; padding:12px 22px; font-weight:800; font-size:14px;
      cursor:pointer; transition:transform .12s, box-shadow .12s;
      box-shadow:0 10px 22px -8px rgba(56,189,248,.6);
    }
    button:hover { transform:translateY(-1px); box-shadow:0 14px 26px -8px rgba(56,189,248,.7); }
    button:active { transform:translateY(0); }
    .btn-upload { background:linear-gradient(135deg,#6366f1,#818cf8); color:#eef2ff;
                  box-shadow:0 10px 22px -8px rgba(99,102,241,.6); }
    .btn-calc { background:linear-gradient(135deg,var(--green),#10b981); color:#04231a;
                box-shadow:0 10px 22px -8px rgba(52,211,153,.6); width:100%; margin-top:4px;
                font-size:15px; padding:14px 22px; }
    .btn-apply { background:linear-gradient(135deg,#f59e0b,#d97706); color:#051426;
                 box-shadow:0 10px 22px -8px rgba(245,158,11,.6); }
    .ok { background:rgba(6,95,70,.30); border:1px solid #10b981; color:#d1fae5;
          padding:14px 16px; border-radius:12px; }
    .fileinfo { margin-top:7px; font-size:12px; color:var(--green); font-weight:600; }

    .filters { width:100%; margin-top:15px; }
    .filtergrid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; }
    .filterbox { border:1px solid var(--line2); border-radius:12px;
                 background:#0a1122; padding:12px; }
    .flabel { display:flex; align-items:center; gap:8px; font-size:11px;
              text-transform:uppercase; letter-spacing:.5px; color:var(--muted);
              font-weight:700; margin-bottom:8px; }
    .fsearch { width:100%; background:#0b1220; border:1px solid var(--line2);
               color:var(--ink); border-radius:8px; padding:8px 10px; font-size:13px;
               margin-bottom:8px; }
    .fsearch:focus { outline:0; border-color:var(--accent);
                     box-shadow:0 0 0 3px rgba(56,189,248,.18); }
    .fselectall { border-bottom:1px solid var(--line2); padding-bottom:8px; margin-bottom:8px; }
    .fselectall label { display:flex; align-items:center; gap:8px; font-size:12px;
                        font-weight:700; color:var(--accent); cursor:pointer;
                        text-transform:none; }
    .flist { max-height:160px; overflow:auto; display:flex; flex-direction:column; gap:6px; }
    .flist .fitem { display:flex; align-items:center; gap:8px; font-size:12px; cursor:pointer; }
    .flist .fitem span { color:#9fc7ff; word-break:break-all; }
    .fhint { color:var(--muted); font-size:12px; padding:6px 2px; }
    input[type=checkbox] { accent-color:var(--accent); width:15px; height:15px; cursor:pointer; }
    
    .filter-actions { display:flex; gap:12px; margin-top:10px; flex-wrap:wrap; }
    .filter-qs { display:flex; align-items:center; gap:8px; color:var(--muted);
                 font-size:13px; font-weight:600; }
    .filter-qs input[type=checkbox] { width:18px; height:18px; }

    .err { background:rgba(127,29,29,.35); border:1px solid #7f1d1d; color:#fecaca;
           padding:14px 16px; border-radius:12px; }

    .bar { display:flex; justify-content:space-between; align-items:center;
           flex-wrap:wrap; gap:12px; margin-bottom:16px; }
    .bar h2 { margin:0; font-size:19px; font-weight:700; }
    .range { color:var(--muted); font-weight:500; font-size:14px; }
    .dl {
      display:inline-flex; align-items:center; gap:7px; text-decoration:none;
      background:linear-gradient(135deg,var(--green),#10b981); color:#04231a;
      padding:11px 18px; border-radius:10px; font-weight:800; font-size:14px;
      box-shadow:0 10px 22px -10px rgba(52,211,153,.7); transition:transform .12s;
    }
    .dl:hover { transform:translateY(-1px); }

    .section-title { font-size:15px; font-weight:700; color:var(--muted);
                     text-transform:uppercase; letter-spacing:.6px; margin:4px 2px 14px; }

    table { border-collapse:collapse; width:100%; font-size:13px; }
    th, td { padding:9px 12px; border-bottom:1px solid var(--line); text-align:right;
             white-space:nowrap; font-variant-numeric:tabular-nums; }
    th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) { text-align:left; }
    thead th { color:var(--muted); font-weight:700; position:sticky; top:0;
               background:var(--card2); font-size:11px; text-transform:uppercase;
               letter-spacing:.5px; border-bottom:1px solid var(--line2); }
    tbody tr { transition:background .12s; }
    tbody tr:hover { background:rgba(56,189,248,.06); }
    tr.total td { font-weight:800; border-top:2px solid var(--accent);
                  background:rgba(56,189,248,.07); }
    .neg { color:var(--red); font-weight:600; } .pos { color:var(--green); font-weight:600; }

    .sumwrap { background:#ffffff; color:#0f172a; border-radius:12px; padding:2px;
               overflow:auto; box-shadow:0 10px 30px -18px rgba(0,0,0,.6); }
    .sumtbl { border-collapse:collapse; width:100%; font-size:13px; color:#0f172a; }
    .sumtbl th, .sumtbl td { padding:8px 12px; border-bottom:1px solid #e6ebf2;
                             text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }
    .sumtbl th:first-child, .sumtbl td:first-child,
    .sumtbl th:nth-child(2), .sumtbl td:nth-child(2) { text-align:left; }
    .sumtbl thead th { background:#1f2a44; color:#fff; position:sticky; top:0;
                       font-size:11px; text-transform:uppercase; letter-spacing:.5px; }
    .sumtbl tr.grand td { background:#e2e8f0; font-weight:800;
                          border-top:2px solid #94a3b8; }

    details {
      background:linear-gradient(180deg,var(--card),var(--bg2));
      border:1px solid var(--line); border-radius:13px; margin-bottom:13px; overflow:hidden;
    }
    details[open] { box-shadow:0 18px 40px -26px rgba(0,0,0,.8); }
    summary { cursor:pointer; padding:14px 18px; font-weight:700; list-style:none;
              display:flex; align-items:center; gap:10px; transition:background .12s; }
    summary::-webkit-details-marker { display:none; }
    summary::before { content:"▸"; color:var(--muted); transition:transform .18s; font-size:12px; }
    details[open] summary::before { transform:rotate(90deg); }
    summary:hover { background:rgba(255,255,255,.03); }
    .count { margin-left:auto; color:var(--muted); font-weight:500; font-size:13px; }
    .scroll { max-height:440px; overflow:auto; padding:0 10px 10px; }
    .badge { font-size:10.5px; padding:3px 10px; border-radius:99px; font-weight:800;
             letter-spacing:.5px; }
    .badge.int { background:rgba(14,116,144,.35); color:#67e8f9; border:1px solid #0e7490; }
    .badge.pos { background:rgba(147,51,234,.30); color:#e9d5ff; border:1px solid #9333ea; }
    
    .filter-badge { font-size:11px; color:var(--muted); background:var(--bg2);
                    padding:4px 12px; border-radius:20px; border:1px solid var(--line2); }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="head">
      <div class="logo">📊</div>
      <h1>INT / POS Algo Calculator</h1>
    </div>
    <p class="sub">Upload the data CSV, pick a date range, get the INT/POS split per algo plus a summary.</p>

    <div class="card">
      <form method="post" action="/process" enctype="multipart/form-data" id="mainForm">
        <div class="field">
          <label>Data CSV</label>
          <input type="file" name="file" accept=".csv">
          {% if data_name %}<div class="fileinfo">📎 Loaded: <b>{{ data_name }}</b> — choose a new file only to replace it.</div>{% endif %}
        </div>
        <div class="field">
          <label>Date from (optional)</label>
          <input type="date" name="date_from" value="{{ date_from or '' }}">
        </div>
        <div class="field">
          <label>Date to (optional)</label>
          <input type="date" name="date_to" value="{{ date_to or '' }}">
        </div>
        <button type="submit" name="action" value="upload" class="btn-upload">⬆️ Upload &amp; load filters</button>

        <input type="hidden" name="data_token" id="data_token" value="{{ data_token }}">

        <!-- Filters Section -->
        <div class="filters">
          <div class="filtergrid">
            <!-- User ID Filter -->
            <div class="filterbox">
              <div class="flabel">👤 User ID</div>
              <input type="text" class="fsearch" placeholder="Search User IDs..." 
                     oninput="filterList(this, 'user-list')">
              <div class="fselectall">
                <label><input type="checkbox" onchange="toggleAll(this, 'user_ids')"> Select All Users</label>
              </div>
              <div class="flist" id="user-list">
                {{ users_html|safe }}
              </div>
            </div>

            <!-- Broker Filter -->
            <div class="filterbox">
              <div class="flabel">🏢 Broker</div>
              <input type="text" class="fsearch" placeholder="Search Brokers..." 
                     oninput="filterList(this, 'broker-list')">
              <div class="fselectall">
                <label><input type="checkbox" onchange="toggleAll(this, 'brokers')"> Select All Brokers</label>
              </div>
              <div class="flist" id="broker-list">
                {{ brokers_html|safe }}
              </div>
            </div>

            <!-- Algo Filter -->
            <div class="filterbox">
              <div class="flabel">📊 Algo</div>
              <input type="text" class="fsearch" placeholder="Search Algos..." 
                     oninput="filterList(this, 'algo-list')">
              <div class="fselectall">
                <label><input type="checkbox" onchange="toggleAll(this, 'algos')"> Select All Algos</label>
              </div>
              <div class="flist" id="algo-list">
                {{ algos_html|safe }}
              </div>
            </div>

            <!-- Server Filter -->
            <div class="filterbox">
              <div class="flabel">🖥️ Server</div>
              <input type="text" class="fsearch" placeholder="Search Servers..." 
                     oninput="filterList(this, 'server-list')">
              <div class="fselectall">
                <label><input type="checkbox" onchange="toggleAll(this, 'servers')"> Select All Servers</label>
              </div>
              <div class="flist" id="server-list">
                {{ servers_html|safe }}
              </div>
            </div>
          </div>
            <button type="submit" name="action" value="apply_filters" class="btn-apply">
              🔍 Apply Filters
            </button>
            <span class="filter-badge" id="filterCount">
              {% if filter_summary %}{{ filter_summary }}{% else %}No filters applied{% endif %}
            </span>
          </div>
        </div>

        <button type="submit" name="action" value="calculate" class="btn-calc">▶️ Calculate</button>
      </form>
    </div>

    {% if notice %}<div class="card ok">{{ notice }}</div>{% endif %}

    {% if error %}<div class="card err">⚠️ {{ error }}</div>{% endif %}

    {% if summary_html %}
      <div class="card">
        <div class="bar">
          <h2>Summary <span class="range">— {{ range_label }}</span></h2>
          <a class="dl" href="/download/{{ token }}">⬇️ Download Excel (all sheets)</a>
        </div>
        {{ summary_html|safe }}
      </div>

      <p class="section-title">Per algo / type</p>
      {% for blk in blocks %}
        <details>
          <summary>Algo {{ blk.algo }}
            <span class="badge {{ 'int' if blk.type=='INT' else 'pos' }}">{{ blk.type }}</span>
            <span class="count">{{ blk.count }} accounts</span></summary>
          <div class="scroll">{{ blk.html|safe }}</div>
        </details>
      {% endfor %}
    {% endif %}
  </div>

  <script>
    function filterList(input, listId) {
      const filter = input.value.toLowerCase();
      const list = document.getElementById(listId);
      const items = list.getElementsByClassName('fitem');
      
      for (let item of items) {
        const text = item.textContent.toLowerCase();
        if (text.includes(filter)) {
          item.style.display = 'flex';
        } else {
          item.style.display = 'none';
        }
      }
    }

    function toggleAll(checkbox, name) {
      const checkboxes = document.querySelectorAll(`input[name="${name}"]`);
      for (let cb of checkboxes) {
        cb.checked = checkbox.checked;
      }
      updateFilterCount();
    }

    function updateFilterCount() {
      const userChecks = document.querySelectorAll('input[name="user_ids"]:checked');
      const brokerChecks = document.querySelectorAll('input[name="brokers"]:checked');
      const algoChecks = document.querySelectorAll('input[name="algos"]:checked');
      const serverChecks = document.querySelectorAll('input[name="servers"]:checked');
      const qsChecked = document.getElementById('qs_filter').checked;
      
      let parts = [];
      if (userChecks.length > 0) parts.push(`👤 ${userChecks.length}`);
      if (brokerChecks.length > 0) parts.push(`🏢 ${brokerChecks.length}`);
      if (algoChecks.length > 0) parts.push(`📊 ${algoChecks.length}`);
      if (serverChecks.length > 0) parts.push(`🖥️ ${serverChecks.length}`);
      if (qsChecked) parts.push('⭐ QS');
      
      const el = document.getElementById('filterCount');
      if (parts.length > 0) {
        el.textContent = 'Filters: ' + parts.join(' | ');
      } else {
        el.textContent = 'No filters applied';
      }
    }

    // Update filter count on any checkbox change
    document.addEventListener('DOMContentLoaded', function() {
      const checkboxes = document.querySelectorAll('input[type="checkbox"]');
      for (let cb of checkboxes) {
        cb.addEventListener('change', updateFilterCount);
      }
      updateFilterCount();
    });
  </script>
</body>
</html>
"""


def _fmt_money(x):
    try:
        return f"{x:,.0f}"
    except (ValueError, TypeError):
        return x


def _fmt_pct(x):
    try:
        cls = "neg" if x < 0 else "pos"
        return f'<span class="{cls}">{x:.2f}%</span>'
    except (ValueError, TypeError):
        return x


def _table_html(df: pd.DataFrame, money_cols, pct_cols, total_label_col=None) -> str:
    d = df.copy()
    for c in money_cols:
        if c in d.columns:
            d[c] = d[c].map(_fmt_money)
    for c in pct_cols:
        if c in d.columns:
            d[c] = d[c].map(_fmt_pct)
    html = d.to_html(index=False, escape=False, border=0)
    if total_label_col:
        html = html.replace("<tr>\n      <td>TOTAL", '<tr class="total">\n      <td>TOTAL')
    return html


def _render_summary(rollup: pd.DataFrame) -> str:
    """Algo-banded summary with INT/POS rows + subtotals + grand rows."""
    if rollup.empty:
        return ""
    g = (rollup.groupby(["Algo", "Type"], as_index=False)
         .agg(Accounts=("Alias", "count"), AUM=("AUM", "sum"), MTM=("Total MTM", "sum")))
    g = g.sort_values(["Algo", "Type"]).reset_index(drop=True)

    def money(x):
        try:
            return f"{x:,.0f}"
        except (ValueError, TypeError):
            return ""

    def signed(x):
        try:
            col = "#c0392b" if x < 0 else "#178f52"
            return f'<span style="color:{col};font-weight:600">{x:,.0f}</span>'
        except (ValueError, TypeError):
            return ""

    def pct(aum, mtm):
        if not aum:
            return ""
        v = mtm / aum * 100
        col = "#c0392b" if v < 0 else "#178f52"
        return f'<span style="color:{col};font-weight:600">{v:.2f}%</span>'

    def cells(algo, typ, acc, aum, mtm):
        return (f"<td>{algo}</td><td>{typ}</td><td>{int(acc):,}</td>"
                f"<td>{money(aum)}</td><td>{signed(mtm)}</td><td>{pct(aum, mtm)}</td>")

    out = ["<div class='sumwrap'><table class='sumtbl'><thead><tr>"
           "<th>Algo</th><th>Type</th><th>Accounts</th>"
           "<th>Total AUM</th><th>Total MTM</th><th>MTM%</th></tr></thead><tbody>"]

    for i, algo in enumerate(dict.fromkeys(g["Algo"].tolist())):
        light, mid = ALGO_BANDS_WEB[i % len(ALGO_BANDS_WEB)]
        block = g[g["Algo"] == algo]
        for _, row in block.iterrows():
            out.append(f"<tr style='background:{light}'>"
                       f"{cells(algo, row['Type'], row['Accounts'], row['AUM'], row['MTM'])}</tr>")
        out.append(f"<tr style='background:{mid};font-weight:700'>"
                   f"{cells(algo, 'Total', block['Accounts'].sum(), block['AUM'].sum(), block['MTM'].sum())}</tr>")

    def grand(label, sub):
        return (f"<tr class='grand'><td colspan='2'>{label}</td>"
                f"<td>{int(sub['Accounts'].sum()):,}</td><td>{money(sub['AUM'].sum())}</td>"
                f"<td>{signed(sub['MTM'].sum())}</td><td>{pct(sub['AUM'].sum(), sub['MTM'].sum())}</td></tr>")

    out.append(grand("TOTAL INT", g[g["Type"] == "INT"]))
    out.append(grand("TOTAL POS", g[g["Type"] == "POS"]))
    out.append(grand("TOTAL OVERALL", g))
    out.append("</tbody></table></div>")
    return "".join(out)


def _get_filtered_df(df: pd.DataFrame, form_data) -> tuple[pd.DataFrame, str, dict]:
    """Apply filters to the dataframe based on form data."""
    filtered_df = df.copy()
    filter_summary_parts = []
    filter_state = {}

    # Get selected values from form
    user_ids = form_data.getlist('user_ids')
    brokers = form_data.getlist('brokers')
    algos = form_data.getlist('algos')
    servers = form_data.getlist('servers')
    qs_filter = form_data.get('qs_filter') == 'on'

    # Apply filters
    if user_ids:
        filtered_df = filtered_df[filtered_df['user_id'].astype(str).isin(user_ids)]
        filter_summary_parts.append(f"👤 {len(user_ids)} users")
        filter_state['user_ids'] = user_ids

    if brokers:
        filtered_df = filtered_df[filtered_df['broker'].astype(str).isin(brokers)]
        filter_summary_parts.append(f"🏢 {len(brokers)} brokers")
        filter_state['brokers'] = brokers

    if algos:
        filtered_df = filtered_df[filtered_df['algo'].astype(str).isin(algos)]
        filter_summary_parts.append(f"📊 {len(algos)} algos")
        filter_state['algos'] = algos

    if servers:
        filtered_df = filtered_df[filtered_df['server'].astype(str).isin(servers)]
        filter_summary_parts.append(f"🖥️ {len(servers)} servers")
        filter_state['servers'] = servers

    if qs_filter:
        # Filter for aliases containing "QS"
        filtered_df = filtered_df[filtered_df['alias'].astype(str).str.contains('QS', case=False, na=False)]
        filter_summary_parts.append("⭐ QS")
        filter_state['qs_filter'] = True

    filter_summary = " | ".join(filter_summary_parts) if filter_summary_parts else "No filters applied"
    
    return filtered_df, filter_summary, filter_state


_PAGE_DEFAULTS = dict(
    summary_html=None, error=None, notice=None, blocks=[], token=None, range_label="",
    date_from="", date_to="", data_token="", data_name="",
    users_html="", brokers_html="", servers_html="", algos_html="",
    filter_summary="No filters applied", qs_filter=False,
)


def _page(**kw):
    return render_template_string(PAGE, **{**_PAGE_DEFAULTS, **kw})


@app.route("/")
def index():
    return _page()


@app.route("/process", methods=["POST"])
def process():
    # Get the CSV bytes
    f = request.files.get("file")
    data_token = request.form.get("data_token") or ""
    
    if f and f.filename:
        raw = f.read()
        data_name = f.filename
        data_token = uuid.uuid4().hex
        _DATA_CACHE[data_token] = (data_name, raw)
        if len(_DATA_CACHE) > 5:
            for k in list(_DATA_CACHE)[:-5]:
                _DATA_CACHE.pop(k, None)
    elif data_token and data_token in _DATA_CACHE:
        data_name, raw = _DATA_CACHE[data_token]
    else:
        return _page(error="Please choose a CSV file.")

    try:
        df = load_csv(raw)
    except (pd.errors.EmptyDataError, UnicodeDecodeError, ValueError, OSError) as e:
        return _page(error=f"Could not read CSV: {e}")

    opts_ctx = dict(data_token=data_token, data_name=data_name)

    # Get filter values for display
    selected_users = request.form.getlist('user_ids')
    selected_brokers = request.form.getlist('brokers')
    selected_algos = request.form.getlist('algos')
    selected_servers = request.form.getlist('servers')
    qs_checked = request.form.get('qs_filter') == 'on'

    # Prepare filter HTML with selected states
    user_options = _unique_options(df, "user_id")
    broker_options = _unique_options(df, "broker")
    algo_options = _unique_options(df, "algo")
    server_options = _unique_options(df, "server")

    users_html = _filter_items_html(user_options, selected_users, "user_ids", "user")
    brokers_html = _filter_items_html(broker_options, selected_brokers, "brokers", "broker")
    algos_html = _filter_items_html(algo_options, selected_algos, "algos", "algo")
    servers_html = _filter_items_html(server_options, selected_servers, "servers", "server")

    # Handle Upload action
    if request.form.get("action") == "upload":
        if df["date"].notna().sum() > 0:
            df_from = request.form.get("date_from") or str(df["date"].min().date())
            df_to = request.form.get("date_to") or str(df["date"].max().date())
        else:
            df_from = request.form.get("date_from") or ""
            df_to = request.form.get("date_to") or ""
        notice = f"✓ Loaded {len(df):,} rows. Select dates and click Calculate."
        return _page(
            date_from=df_from, date_to=df_to, notice=notice,
            users_html=users_html, brokers_html=brokers_html,
            algos_html=algos_html, servers_html=servers_html,
            qs_filter=qs_checked,
            **opts_ctx
        )

    # Handle Apply Filters action
    if request.form.get("action") == "apply_filters":
        # Apply filters to show preview
        filtered_df, filter_summary, _ = _get_filtered_df(df, request.form)
        
        # Get date range
        if df["date"].notna().sum() > 0:
            min_d, max_d = df["date"].min().date(), df["date"].max().date()
            df_from = request.form.get("date_from") or str(min_d)
            df_to = request.form.get("date_to") or str(max_d)
        else:
            df_from = request.form.get("date_from") or ""
            df_to = request.form.get("date_to") or ""
        
        # Show how many rows after filtering
        notice = f"✓ Filters applied: {len(filtered_df):,} rows after filtering (out of {len(df):,} total). Click Calculate to see results."
        
        # Store filter state for calculation
        _FILTER_CACHE[data_token] = (filtered_df, dict(request.form))
        
        return _page(
            date_from=df_from, date_to=df_to, notice=notice,
            users_html=users_html, brokers_html=brokers_html,
            algos_html=algos_html, servers_html=servers_html,
            filter_summary=filter_summary,
            qs_filter=qs_checked,
            **opts_ctx
        )

    # Handle Calculate action
    if df["date"].notna().sum() == 0:
        return _page(error="Could not parse the 'date' column.", **opts_ctx)

    # Check if we have a cached filtered dataframe from Apply Filters
    if data_token in _FILTER_CACHE:
        cached_df, filter_state = _FILTER_CACHE[data_token]
        # Use the cached filtered dataframe
        df = cached_df
        # Get filter summary
        _, filter_summary, _ = _get_filtered_df(cached_df, request.form)
    else:
        # No filters applied, use full dataframe
        filter_summary = "No filters applied"

    min_d, max_d = df["date"].min().date(), df["date"].max().date()
    df_from = request.form.get("date_from") or str(min_d)
    df_to = request.form.get("date_to") or str(max_d)
    
    try:
        d_from = pd.to_datetime(df_from).date()
        d_to = pd.to_datetime(df_to).date()
    except (ValueError, TypeError):
        return _page(error="Invalid date(s).", **opts_ctx)
    
    if d_from > d_to:
        return _page(
            error="'Date from' must be on or before 'Date to'.",
            date_from=df_from, date_to=df_to,
            users_html=users_html, brokers_html=brokers_html,
            algos_html=algos_html, servers_html=servers_html,
            qs_filter=qs_checked,
            **opts_ctx
        )

    mask = (df["date"].dt.date >= d_from) & (df["date"].dt.date <= d_to)
    df_range = df.loc[mask].copy()
    
    if df_range.empty:
        return _page(
            error="No rows in the selected date range.",
            date_from=df_from, date_to=df_to,
            users_html=users_html, brokers_html=brokers_html,
            algos_html=algos_html, servers_html=servers_html,
            qs_filter=qs_checked,
            **opts_ctx
        )

    rollup = build_account_rollup(df_range)

    # Per algo/type blocks
    blocks = []
    detail_cols = ["Alias", "UserID", "AUM", "Total MTM", "MTM%", "Trading Days", "Broker", "Server"]
    for (algo, acct_type), g in rollup.groupby(["Algo", "Type"], sort=True):
        blocks.append({
            "algo": algo, "type": acct_type, "count": len(g),
            "html": _table_html(g[detail_cols], ["AUM", "Total MTM"], ["MTM%"]),
        })

    summary_html = _render_summary(rollup)

    token = uuid.uuid4().hex
    _EXCEL_CACHE[token] = (f"int_pos_{d_from}_{d_to}.xlsx", build_excel(rollup, df_range))
    if len(_EXCEL_CACHE) > 20:
        for k in list(_EXCEL_CACHE)[:-20]:
            _EXCEL_CACHE.pop(k, None)

    return _page(
        summary_html=summary_html, blocks=blocks, token=token,
        range_label=f"{d_from} to {d_to}",
        date_from=df_from, date_to=df_to,
        users_html=users_html, brokers_html=brokers_html,
        algos_html=algos_html, servers_html=servers_html,
        filter_summary=filter_summary,
        qs_filter=qs_checked,
        **opts_ctx
    )


@app.route("/download/<token>")
def download(token):
    item = _EXCEL_CACHE.get(token)
    if not item:
        abort(404, "Download expired — please re-run the calculation.")
    fname, data = item
    return send_file(
        io.BytesIO(data), as_attachment=True, download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)