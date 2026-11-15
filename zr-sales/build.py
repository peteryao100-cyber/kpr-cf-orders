#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build ZR Full Sales Analysis v4 - 完全由筛选驱动的互动页面
第四稿：整页由顶部筛选驱动，不是静态报告外挂查询器
"""
from __future__ import annotations
import json
import math
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
import numpy as np

BASE = Path(__file__).parent
DATA = BASE / "data"
OUT = BASE.parent  # Output HTML to repo root

F1 = DATA / "ZR销售月统计_全部_202401_202512.xlsx"
F2 = DATA / "ZR销售月统计_全部_202601_202610.xlsx"
F3 = DATA / "ZR销售月统计_20261001-20261115.xlsx"

FX_FILES = {
    2024: DATA / "fx_2024.json",
    2025: DATA / "fx_2025.json",
    2026: DATA / "fx_2026.json",
}

LV_CUSTOMERS = {
    "上海绿安全(华南)": "华南",
    "上海绿安全(华东)": "华东",
    "上海绿安全(华北)": "华北",
    "上海绿安全外贸": "外贸",
    "上海外贸绿安全": "外贸",
}

PART_ORD = {"上": 1, "中": 2, "下": 3}


def strip_cell(x):
    return x.strip() if isinstance(x, str) else x


def load_fx_rates() -> dict[str, float]:
    """Load monthly average FX rates from JSON files."""
    rates = {}
    for year, fpath in FX_FILES.items():
        data = json.loads(fpath.read_text(encoding="utf-8"))
        by_month = defaultdict(list)
        
        start_date = f"{year}-01-01"
        end_date = f"{year}-12-31" if year < 2026 else "2026-11-13"
        
        for dt_str, rec in data["rates"].items():
            if dt_str < start_date or dt_str > end_date:
                continue
            if not dt_str.startswith(str(year)):
                continue
            ym = dt_str[:7]
            by_month[ym].append(rec["CNY"])
        
        for ym, vals in by_month.items():
            rates[ym] = sum(vals) / len(vals)
    
    return rates


def load_and_merge_data() -> pd.DataFrame:
    """Load and merge the three Excel files according to the merge rules."""
    df1 = pd.read_excel(F1)
    df2 = pd.read_excel(F2)
    df3 = pd.read_excel(F3)
    
    df2_keep = df2[~df2["日期"].astype(str).str.startswith("2026-10")].copy()
    
    df = pd.concat([
        df1.assign(_src="file1"),
        df2_keep.assign(_src="file2"),
        df3.assign(_src="file3"),
    ], ignore_index=True)
    
    for c in ["客户简称", "国别/通路", "品牌", "客户型号", "鞋型", "类别", "币别", "客户编号"]:
        if c in df.columns:
            df[c] = df[c].map(strip_cell)
    
    df["year"] = df["日期"].str[:4].astype(int)
    df["month"] = df["日期"].str[5:7].astype(int)
    df["part"] = df["日期"].str[7]
    df["ym"] = df["日期"].str[:7]
    df["xun_ord"] = df["year"] * 1000 + df["month"] * 10 + df["part"].map(PART_ORD)
    df["same"] = df["xun_ord"] <= df["year"] * 1000 + 111
    df["qty"] = df["销售数量"].astype(int)
    df["unit_price"] = df["单价"].astype(float)
    df["fee"] = df["服务费"].fillna(0).astype(float)
    df["amt_native"] = (df["unit_price"] + df["fee"]) * df["qty"]
    
    return df


def m3_bucket(xing, xinghao) -> str:
    """Classify 3M items into subcategories."""
    t = f"{xinghao or ''} {xing or ''}"
    if any(k in t for k in ["耳塞", "耳罩"]):
        return "耳塞"
    if any(k in t for k in ["口罩", "防护面具", "防毒面具", "防毒面罩", "半面罩"]):
        return "口罩"
    return "其他"


def classify_row(r) -> tuple[str, str, bool, str]:
    """
    Classify a row into unit, subunit, is_shoe, and filter_unit.
    Returns: (unit, subunit, is_shoe, filter_unit)
    filter_unit is a fine-grained classification for filtering.
    """
    typ = r["类别"] or ""
    brand = r["品牌"] or ""
    name = r["客户简称"] or ""
    xing = r["鞋型"] or ""
    xinghao = r["客户型号"] or ""
    
    if name in LV_CUSTOMERS:
        region = LV_CUSTOMERS[name]
        if typ in ("MIDORI", "内销"):
            fu = f"绿安全{region}" if region != "外贸" else "绿安全外贸"
            return "绿安全", "成品鞋", True, fu
        if brand == "鞋垫" or typ == "其它":
            fu = f"绿安全{region}" if region != "外贸" else "绿安全外贸"
            return "绿安全", "鞋垫", False, fu
        fu = f"绿安全{region}" if region != "外贸" else "绿安全外贸"
        return "绿安全", f"其他({typ})", False, fu
    
    if typ == "内销":
        if brand == "半成品-KPR" or "半成品" in xing or "Upper" in str(xinghao):
            return "OEM鞋面/半成品", "半成品-KPR（原记内销）", False, "OEM鞋面/半成品"
        return "内销", "成品鞋", True, "内销"
    
    if typ == "外销":
        if brand == "半成品-KPR" or "半成品" in xing or "Shoe Uppers" in str(xinghao) or "Upper" in str(xinghao):
            return "OEM鞋面/半成品", "半成品-KPR（原记外销）", False, "OEM鞋面/半成品"
        return "外销", "成品鞋", True, "外销"
    
    if typ == "OEM":
        return "OEM鞋面/半成品", "typ=OEM（CF鞋面套装）", False, "OEM鞋面/半成品"
    
    if typ == "钢头":
        return "钢头", "", False, "钢头"
    
    if typ == "橡胶":
        return "橡胶底", "", False, "橡胶底"
    
    if typ == "3M":
        sub = m3_bucket(xing, xinghao)
        return "3M", sub, False, f"3M{sub}"
    
    if typ == "MIDORI":
        return "X", "其他成品（非上海绿安全）", True, "X"
    
    if typ == "其它":
        if brand == "外购鞋":
            return "X", "外购鞋", True, "X"
        if brand == "KRP":
            return "X", "其他成品（品牌KRP，售CF）", True, "X"
        if brand == "材料":
            return "材料", "", False, "材料"
        if brand == "鞋垫":
            return "鞋垫", "", False, "鞋垫"
        if brand == "其它劳保":
            return "其它劳保", "", False, "其它劳保"
        if "A3-089" in xing:
            return "X", "其他成品（其它未标品牌，鞋型A3-089）", True, "X"
        return "其它未标品牌（材料/配件）", "", False, "其它未标品牌（材料/配件）"
    
    return "未分类", typ, False, "未分类"


def apply_fx_and_classify(df: pd.DataFrame, rates: dict[str, float]) -> pd.DataFrame:
    """Apply FX rates and classify each row."""
    df = df.copy()
    
    cls = df.apply(classify_row, axis=1, result_type="expand")
    df["unit"] = cls[0]
    df["subunit"] = cls[1]
    df["is_shoe"] = cls[2]
    df["filter_unit"] = cls[3]
    
    df["cur"] = df["币别"].fillna("")
    df["fx"] = df["ym"].map(rates)
    df["amt_flag"] = ""
    
    empty = df["cur"].eq("")
    other = ~df["cur"].isin(["RMB", "USD", ""])
    df.loc[empty, "amt_flag"] = "币别为空"
    df.loc[other, "amt_flag"] = "未知币别"
    
    df["amt_rmb"] = np.nan
    df["amt_usd"] = np.nan
    df["amt_usd_rmb"] = np.nan
    
    rmb = df["cur"].eq("RMB")
    usd = df["cur"].eq("USD")
    df.loc[rmb, "amt_rmb"] = df.loc[rmb, "amt_native"]
    df.loc[usd, "amt_usd"] = df.loc[usd, "amt_native"]
    
    has_fx = usd & df["fx"].notna()
    df.loc[has_fx, "amt_usd_rmb"] = df.loc[has_fx, "amt_native"] * df.loc[has_fx, "fx"]
    
    df["amt_rmb_total"] = df["amt_rmb"].fillna(0) + df["amt_usd_rmb"].fillna(0)
    df.loc[empty | other, "amt_rmb_total"] = 0
    
    df["amt_fen"] = (df["amt_rmb_total"] * 100).round().astype(int)
    
    df["include_amt"] = (df["amt_flag"] == "").astype(int)
    
    return df


FILTER_UNITS = [
    "内销",
    "外销",
    "绿安全华南",
    "绿安全华东",
    "绿安全华北",
    "绿安全外贸",
    "X",
    "OEM鞋面/半成品",
    "钢头",
    "橡胶底",
    "3M口罩",
    "3M耳塞",
    "3M其他",
    "材料",
    "鞋垫",
    "其它劳保",
    "其它未标品牌（材料/配件）",
]
FU_INDEX = {n: i for i, n in enumerate(FILTER_UNITS)}


def pack_data_for_js(df: pd.DataFrame, rates: dict[str, float]) -> dict:
    """Pack all data into a format suitable for JavaScript filtering."""
    cust_list = sorted({str(x) if pd.notna(x) and str(x) != "nan" else "（空）" 
                        for x in df["客户简称"].tolist()})
    style_list = sorted({str(x) if pd.notna(x) and str(x) != "nan" else "（空）" 
                         for x in df["鞋型"].tolist()})
    ch_list = sorted({str(x) if pd.notna(x) and str(x) != "nan" else "（空）" 
                      for x in df["国别/通路"].tolist()})
    brand_list = sorted({str(x) if pd.notna(x) and str(x) != "nan" else "（空）" 
                         for x in df["品牌"].tolist()})
    
    cust_idx = {n: i for i, n in enumerate(cust_list)}
    style_idx = {n: i for i, n in enumerate(style_list)}
    ch_idx = {n: i for i, n in enumerate(ch_list)}
    brand_idx = {n: i for i, n in enumerate(brand_list)}
    
    rows = []
    for idx, r in df.iterrows():
        fu = r["filter_unit"]
        if fu not in FU_INDEX:
            print(f"Warning: unknown filter_unit '{fu}' for row {idx}")
            continue
        
        cust = str(r["客户简称"]) if pd.notna(r["客户简称"]) and str(r["客户简称"]) != "nan" else "（空）"
        style = str(r["鞋型"]) if pd.notna(r["鞋型"]) and str(r["鞋型"]) != "nan" else "（空）"
        ch = str(r["国别/通路"]) if pd.notna(r["国别/通路"]) and str(r["国别/通路"]) != "nan" else "（空）"
        brand = str(r["品牌"]) if pd.notna(r["品牌"]) and str(r["品牌"]) != "nan" else "（空）"
        
        rows.append([
            FU_INDEX[fu],
            int(r["year"]),
            int(r["month"]),
            PART_ORD[r["part"]],
            1 if r["same"] else 0,
            cust_idx[cust],
            style_idx[style],
            ch_idx[ch],
            brand_idx[brand],
            int(r["qty"]),
            int(r["amt_fen"]),
            1 if r["is_shoe"] else 0,
            int(r["include_amt"]),
        ])
    
    fx_monthly = []
    for ym in sorted(rates.keys()):
        fx_monthly.append({"ym": ym, "rate": rates[ym]})
    
    return {
        "fu": FILTER_UNITS,
        "cust": cust_list,
        "style": style_list,
        "ch": ch_list,
        "brand": brand_list,
        "rows": rows,
        "fx": fx_monthly,
    }


def generate_html(data: dict) -> str:
    """Generate the complete interactive HTML page."""
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    data_json = data_json.replace("<", "\\u003c").replace(">", "\\u003e")
    
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>尊荣（ZR）完整销售分析 · 第四稿（全页筛选驱动）</title>
<style>
:root{{--ink:#1a2330;--muted:#5c6770;--line:#e6eaee;--bg:#f4f6f8;--card:#fff;--blue:#1e4d7b;--green:#1f6b4a;--up:#1f6b4a;--down:#b42318;--warn:#c05621}}
*{{box-sizing:border-box}}
html{{scroll-behavior:smooth}}
body{{margin:0;font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Noto Sans SC","Segoe UI",sans-serif;color:var(--ink);background:var(--bg);line-height:1.55}}
header{{background:linear-gradient(135deg,#143654,#1e4d7b);color:#fff;padding:20px 20px 16px}}
header h1{{margin:0 0 4px;font-size:20px;font-weight:700}}
header p{{margin:0;opacity:.88;font-size:12px}}
.sticky-filters{{position:sticky;top:0;z-index:100;background:var(--card);border-bottom:1px solid var(--line);padding:12px 16px;box-shadow:0 2px 8px rgba(0,0,0,.08)}}
.filters{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:8px 12px}}
.filters label{{display:flex;flex-direction:column;gap:3px;font-size:11px;color:var(--muted);font-weight:600}}
.filters select,.filters input{{font-size:16px;padding:8px 10px;border:1px solid var(--line);border-radius:6px;background:#fff;color:var(--ink);min-height:44px;width:100%;font-family:inherit}}
.filter-row{{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px;align-items:center}}
.btn{{cursor:pointer;background:var(--blue);color:#fff;border:0;font-weight:600;padding:10px 14px;border-radius:6px;font-size:14px;min-height:44px}}
.btn-lite{{background:#eef2f6;color:var(--ink)}}
.filter-summary{{font-size:12px;color:var(--muted);flex:1;min-width:200px}}
.wrap{{max-width:1200px;margin:0 auto;padding:16px 16px 48px}}
.note{{background:#fff;border:1px solid var(--line);border-left:4px solid var(--blue);padding:10px 14px;margin:12px 0;font-size:12px;color:var(--muted)}}
.note.warn{{border-left-color:var(--warn)}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:12px 0}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px}}
.card .k{{font-size:11px;color:var(--muted)}}
.card .v{{font-size:20px;font-weight:700;margin:3px 0 2px;letter-spacing:-.02em}}
.card .s{{font-size:11px;color:var(--muted)}}
.chg{{margin-top:4px;font-size:12px;font-weight:600}}
.chg.up{{color:var(--up)}} .chg.down{{color:var(--down)}} .chg.flat{{color:#667}}
.panel{{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:14px 16px;margin:12px 0}}
.panel h2{{margin:0 0 6px;font-size:15px}}
.panel h3{{margin:12px 0 6px;font-size:13px}}
.panel .sub{{color:var(--muted);font-size:11px;margin:0 0 8px}}
.scroll{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th,td{{padding:5px 6px;border-bottom:1px solid var(--line);text-align:left}}
th{{font-size:11px;color:var(--muted);font-weight:600;background:#fafbfc;position:sticky;top:0}}
td.num,th.num{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
tr.sum td{{font-weight:700;background:#f6f8fa}}
tr.hi td{{background:#fff6e8}}
.up{{color:var(--up)!important}} .down{{color:var(--down)!important}}
.tag{{display:inline-block;font-size:9px;background:#fff3e8;color:var(--warn);border-radius:3px;padding:1px 4px;margin-left:3px}}
footer{{font-size:11px;color:var(--muted);margin-top:12px;line-height:1.6;padding:12px 16px;background:var(--card);border:1px solid var(--line);border-radius:8px}}
code{{font-size:11px;background:#f0f3f6;padding:1px 3px;border-radius:2px}}
svg{{max-width:100%;height:auto;display:block}}
.chart-container{{margin:8px 0}}
@media(max-width:800px){{
  header{{padding:16px 12px 12px}}
  header h1{{font-size:17px}}
  .sticky-filters{{position:static;padding:10px 12px}}
  .mobile-sticky{{position:sticky;top:0;z-index:100;background:var(--card);border-bottom:1px solid var(--line);padding:8px 12px;box-shadow:0 2px 8px rgba(0,0,0,.08)}}
  .filters{{grid-template-columns:1fr 1fr}}
  .filters-extra{{display:none}}
  .filters-extra.show{{display:contents}}
  .wrap{{padding:10px 10px 36px}}
  .panel{{padding:10px 12px}}
  .card .v{{font-size:18px}}
  table{{font-size:11px}}
  .btn-expand{{display:inline-block}}
  .filter-row{{flex-wrap:wrap}}
}}
@media(min-width:801px){{
  .mobile-sticky{{display:none}}
  .btn-expand{{display:none}}
  .filters-extra{{display:contents}}
}}
.hidden{{display:none!important}}
#loading{{text-align:center;padding:40px;color:var(--muted)}}
</style>
</head>
<body>
<header>
  <h1>尊荣（ZR）完整销售分析</h1>
  <p>第四稿 · 整页由筛选驱动 · 2024 年 — 2026 年 11 月 15 日</p>
</header>

<div class="mobile-sticky">
  <div class="filter-summary" id="filter-summary-mobile"></div>
</div>

<div class="sticky-filters">
  <div class="filters">
    <label>期间
      <select id="f-period">
        <option value="sp2026" selected>2026 同期（1月–11上）</option>
        <option value="sp2025">2025 同期（1月–11上）</option>
        <option value="sp2024">2024 同期（1月–11上）</option>
        <option value="sp_compare">三年同期对比</option>
        <option value="fy2025">2025 全年</option>
        <option value="fy2024">2024 全年</option>
        <option value="ytd2026">2026 截至 11-15（含不完整11中）</option>
        <option value="all">三年全部（2026 截至 11-15）</option>
      </select>
    </label>
    <label>事业 / 品类
      <select id="f-unit">
        <option value="shoe" selected>成品鞋合计</option>
        <option value="all">全部单元（只加营业额）</option>
        <option value="内销">内销</option>
        <option value="外销">外销</option>
        <option value="绿安全国内">绿安全国内（华南+华东+华北）</option>
        <option value="绿安全华南">绿安全华南</option>
        <option value="绿安全华东">绿安全华东</option>
        <option value="绿安全华北">绿安全华北</option>
        <option value="绿安全外贸">绿安全外贸</option>
        <option value="X">X（其他成品鞋）</option>
        <option value="OEM鞋面/半成品">OEM鞋面 / 半成品</option>
        <option value="钢头">钢头</option>
        <option value="橡胶底">橡胶底</option>
        <option value="3M">3M全部</option>
        <option value="3M口罩">3M口罩</option>
        <option value="3M耳塞">3M耳塞（含耳罩）</option>
        <option value="3M其他">3M其他</option>
        <option value="材料">材料</option>
        <option value="鞋垫">鞋垫（不含绿安全名下）</option>
        <option value="其它劳保">其它劳保</option>
        <option value="其它未标品牌（材料/配件）">其它未标品牌（材料/配件）</option>
      </select>
    </label>
    <div class="filters-extra" id="filters-extra">
      <label>客户 <input id="f-cust" list="dl-cust" placeholder="全部客户" autocomplete="off"/></label>
      <label>鞋型 <input id="f-style" list="dl-style" placeholder="全部鞋型" autocomplete="off"/></label>
      <label>通路/地区
        <select id="f-ch"><option value="">全部</option></select>
      </label>
      <label>品牌
        <select id="f-brand"><option value="">全部</option></select>
      </label>
    </div>
  </div>
  <div class="filter-row">
    <button class="btn btn-lite" id="btn-reset">重置为默认</button>
    <button class="btn btn-lite btn-expand" id="btn-expand">更多筛选 ▼</button>
    <div class="filter-summary" id="filter-summary"></div>
  </div>
</div>
<datalist id="dl-cust"></datalist>
<datalist id="dl-style"></datalist>

<div class="wrap">
  <div id="loading">正在加载数据...</div>
  
  <div id="main-content" class="hidden">
    <div class="note" id="note-box">
      <span id="note-text"></span>
    </div>
    
    <div class="grid" id="kpi-grid"></div>
    
    <div class="panel">
      <h2>构成拆分</h2>
      <p class="sub" id="split-note"></p>
      <div class="scroll"><table id="tbl-split"><thead></thead><tbody></tbody></table></div>
    </div>
    
    <div class="panel">
      <h2>营业额（按年）</h2>
      <div class="chart-container" id="chart-year"></div>
    </div>
    
    <div class="panel">
      <h2>月度走势</h2>
      <div class="chart-container" id="chart-month"></div>
    </div>
    
    <div class="panel">
      <h2>按月明细</h2>
      <div class="scroll"><table id="tbl-month"><thead></thead><tbody></tbody></table></div>
    </div>
    
    <div class="panel">
      <h2>客户（营业额前 20）</h2>
      <p class="sub" id="cust-note">2026 年 10 月起内销客户大量记为 NEWTEX1，解读时请注意。</p>
      <div class="scroll"><table id="tbl-cust"><thead></thead><tbody></tbody></table></div>
    </div>
    
    <div class="panel">
      <h2>鞋型（营业额前 20）</h2>
      <div class="scroll"><table id="tbl-style"><thead></thead><tbody></tbody></table></div>
    </div>
    
    <div class="panel">
      <h2>通路 / 地区</h2>
      <div class="scroll"><table id="tbl-ch"><thead></thead><tbody></tbody></table></div>
    </div>
    
    <div class="panel">
      <h2>品牌</h2>
      <div class="scroll"><table id="tbl-brand"><thead></thead><tbody></tbody></table></div>
    </div>
    
    <footer>
      <strong>口径说明</strong><br>
      • 行金额 =（单价 + 服务费）× 销售数量。美元按 Frankfurter（ECB 参考汇率）当月平均 USDCNY 折人民币。<br>
      • 同期 = 1 月上旬至 11 月上旬（完整旬）。2026-11中只到 11 月 15 日，不计入同期。<br>
      • 13 行（1,074 双，内销 NEWTEX1）币别为空，数量计入、营业额不计入。<br>
      • 成品鞋合计 = 内销 + 外销 + 绿安全国内 + 绿安全外贸 + X。选「全部单元」时只加营业额，不加数量。<br>
      • 绿安全国内 = 华南 + 华东 + 华北，不含外贸。鞋垫单元不含绿安全名下鞋垫（留在绿安全各区）。<br>
      • 3M 的口罩、耳塞、其他件数单位不同，不应加总；营业额可以加总。<br>
      • 2026 年 10 月起内销客户大量塌缩为 NEWTEX1，通路常为「其他」。<br>
      <br>
      <strong>美元 → 人民币月均汇率</strong>（来源：Frankfurter / ECB）<br>
      <span id="fx-table"></span>
    </footer>
  </div>
</div>

<script>
const DATA = {data_json};

const FU = DATA.fu;
const CUST = DATA.cust;
const STYLE = DATA.style;
const CH = DATA.ch;
const BRAND = DATA.brand;
const ROWS = DATA.rows;
const FX = DATA.fx;

const SHOE_FU = new Set(["内销", "外销", "绿安全华南", "绿安全华东", "绿安全华北", "绿安全外贸", "X"]);
const LV_DOM_FU = new Set(["绿安全华南", "绿安全华东", "绿安全华北"]);
const M3_FU = new Set(["3M口罩", "3M耳塞", "3M其他"]);

function getFuIdx(fu) {{ return FU.indexOf(fu); }}

const COLS = {{ FU: 0, YEAR: 1, MONTH: 2, PART: 3, SAME: 4, CUST: 5, STYLE: 6, CH: 7, BRAND: 8, QTY: 9, FEN: 10, SHOE: 11, INC: 12 }};

function filterRows(period, unitSel, custStr, styleStr, chVal, brandVal) {{
  let matchYear, matchSame, matchYtd;
  const years = [];
  if (period === "sp2026") {{ years.push(2026); matchSame = true; }}
  else if (period === "sp2025") {{ years.push(2025); matchSame = true; }}
  else if (period === "sp2024") {{ years.push(2024); matchSame = true; }}
  else if (period === "sp_compare") {{ years.push(2024, 2025, 2026); matchSame = true; }}
  else if (period === "fy2025") {{ years.push(2025); matchSame = false; }}
  else if (period === "fy2024") {{ years.push(2024); matchSame = false; }}
  else if (period === "ytd2026") {{ years.push(2026); matchSame = false; matchYtd = true; }}
  else if (period === "all") {{ years.push(2024, 2025, 2026); matchSame = false; }}
  
  const yearSet = new Set(years);
  
  let fuSet = null;
  if (unitSel === "all") {{ fuSet = null; }}
  else if (unitSel === "shoe") {{ fuSet = SHOE_FU; }}
  else if (unitSel === "绿安全国内") {{ fuSet = LV_DOM_FU; }}
  else if (unitSel === "3M") {{ fuSet = M3_FU; }}
  else {{ fuSet = new Set([unitSel]); }}
  
  const custLower = custStr ? custStr.toLowerCase() : "";
  const styleLower = styleStr ? styleStr.toLowerCase() : "";
  const chIdx = chVal ? CH.indexOf(chVal) : -1;
  const brandIdx = brandVal ? BRAND.indexOf(brandVal) : -1;
  
  const result = [];
  for (const r of ROWS) {{
    if (!yearSet.has(r[COLS.YEAR])) continue;
    if (matchSame && !r[COLS.SAME]) continue;
    if (fuSet && !fuSet.has(FU[r[COLS.FU]])) continue;
    if (custLower && !CUST[r[COLS.CUST]].toLowerCase().includes(custLower)) continue;
    if (styleLower && !STYLE[r[COLS.STYLE]].toLowerCase().includes(styleLower)) continue;
    if (chIdx >= 0 && r[COLS.CH] !== chIdx) continue;
    if (brandIdx >= 0 && r[COLS.BRAND] !== brandIdx) continue;
    result.push(r);
  }}
  return result;
}}

function aggregate(rows) {{
  let qty = 0, fen = 0, rowCount = 0, shoeQty = 0, shoeFen = 0, excQty = 0, excRows = 0;
  let m3Kouzhao = 0, m3Ersai = 0, m3Qita = 0;
  let insoleFen = 0, insoleQty = 0;
  const custSet = new Set(), styleSet = new Set();
  for (const r of rows) {{
    rowCount++;
    qty += r[COLS.QTY];
    custSet.add(r[COLS.CUST]);
    styleSet.add(r[COLS.STYLE]);
    const fu = FU[r[COLS.FU]];
    if (fu === "3M口罩") m3Kouzhao += r[COLS.QTY];
    else if (fu === "3M耳塞") m3Ersai += r[COLS.QTY];
    else if (fu === "3M其他") m3Qita += r[COLS.QTY];
    if (r[COLS.SHOE]) {{
      shoeQty += r[COLS.QTY];
      if (r[COLS.INC]) shoeFen += r[COLS.FEN];
    }} else if (fu.startsWith("绿安全") && !r[COLS.SHOE]) {{
      insoleQty += r[COLS.QTY];
      if (r[COLS.INC]) insoleFen += r[COLS.FEN];
    }}
    if (r[COLS.INC]) {{
      fen += r[COLS.FEN];
    }} else {{
      excQty += r[COLS.QTY];
      excRows++;
    }}
  }}
  return {{ 
    qty, fen, rowCount, shoeQty, shoeFen, excQty, excRows, 
    yuan: fen / 100, shoeYuan: shoeFen / 100,
    m3Kouzhao, m3Ersai, m3Qita,
    insoleQty, insoleYuan: insoleFen / 100,
    custCount: custSet.size, styleCount: styleSet.size
  }};
}}

function byYear(rows) {{
  const m = new Map();
  for (const r of rows) {{
    const y = r[COLS.YEAR];
    if (!m.has(y)) m.set(y, []);
    m.get(y).push(r);
  }}
  return m;
}}

function byMonth(rows) {{
  const m = new Map();
  for (const r of rows) {{
    const ym = r[COLS.YEAR] + "-" + String(r[COLS.MONTH]).padStart(2, "0");
    if (!m.has(ym)) m.set(ym, []);
    m.get(ym).push(r);
  }}
  return m;
}}

function byDim(rows, dimIdx) {{
  const m = new Map();
  for (const r of rows) {{
    const k = r[dimIdx];
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }}
  return m;
}}

function byFu(rows) {{
  const m = new Map();
  for (const r of rows) {{
    const fu = FU[r[COLS.FU]];
    if (!m.has(fu)) m.set(fu, []);
    m.get(fu).push(r);
  }}
  return m;
}}

function fmt(n) {{ return n.toLocaleString("zh-CN"); }}
function fmtYuan(n) {{ return Math.round(n).toLocaleString("zh-CN"); }}
function fmtWan(n) {{ return (n / 10000).toLocaleString("zh-CN", {{ minimumFractionDigits: 1, maximumFractionDigits: 1 }}); }}
function fmtYi(n) {{ return (n / 100000000).toLocaleString("zh-CN", {{ minimumFractionDigits: 3, maximumFractionDigits: 3 }}); }}
function pct(newV, oldV) {{ return oldV ? ((newV / oldV - 1) * 100).toFixed(1) + "%" : "—"; }}
function pctCls(newV, oldV) {{ if (!oldV) return ""; const p = newV / oldV - 1; return p > 0.0005 ? "up" : p < -0.0005 ? "down" : ""; }}

function canShowQty(unitSel) {{
  if (unitSel === "all" || unitSel === "3M") return false;
  return true;
}}

function qtyUnit(unitSel) {{
  if (unitSel === "shoe" || unitSel === "内销" || unitSel === "外销" || 
      unitSel.startsWith("绿安全") || unitSel === "X") return "双";
  return "件";
}}

function render() {{
  const period = document.getElementById("f-period").value;
  const unitSel = document.getElementById("f-unit").value;
  const custStr = document.getElementById("f-cust").value.trim();
  const styleStr = document.getElementById("f-style").value.trim();
  const chVal = document.getElementById("f-ch").value;
  const brandVal = document.getElementById("f-brand").value;
  
  const rows = filterRows(period, unitSel, custStr, styleStr, chVal, brandVal);
  const agg = aggregate(rows);
  
  updateFilters(rows, chVal, brandVal);
  updateSummary(rows, agg, period, unitSel, custStr, styleStr, chVal, brandVal);
  updateNote(unitSel, agg);
  updateKPIs(rows, agg, period, unitSel);
  updateSplit(rows, period, unitSel);
  updateYearChart(rows, unitSel);
  updateMonthChart(rows, unitSel);
  updateMonthTable(rows, unitSel);
  updateDimTable("tbl-cust", rows, COLS.CUST, CUST, 20, unitSel);
  updateDimTable("tbl-style", rows, COLS.STYLE, STYLE, 20, unitSel);
  updateDimTable("tbl-ch", rows, COLS.CH, CH, 50, unitSel);
  updateDimTable("tbl-brand", rows, COLS.BRAND, BRAND, 50, unitSel);
  
  document.getElementById("cust-note").style.display = 
    (unitSel === "内销" || unitSel === "shoe" || unitSel === "all") ? "" : "none";
}}

function updateFilters(rows, currentCh, currentBrand) {{
  const chSet = new Set();
  const brandSet = new Set();
  for (const r of ROWS) {{ chSet.add(r[COLS.CH]); brandSet.add(r[COLS.BRAND]); }}
  
  const chSel = document.getElementById("f-ch");
  const brandSel = document.getElementById("f-brand");
  
  const chOpts = ['<option value="">全部</option>'];
  [...chSet].map(i => CH[i]).sort().forEach(v => {{
    const sel = v === currentCh ? " selected" : "";
    chOpts.push(`<option value="${{v}}"${{sel}}>${{v}}</option>`);
  }});
  chSel.innerHTML = chOpts.join("");
  
  const brandOpts = ['<option value="">全部</option>'];
  [...brandSet].map(i => BRAND[i]).sort().forEach(v => {{
    const sel = v === currentBrand ? " selected" : "";
    brandOpts.push(`<option value="${{v}}"${{sel}}>${{v}}</option>`);
  }});
  brandSel.innerHTML = brandOpts.join("");
}}

function getDisplayQty(agg, unitSel) {{
  if (unitSel === "shoe" || unitSel === "绿安全国内") return agg.shoeQty;
  return agg.qty;
}}

function getDisplayYuan(agg, unitSel) {{
  if (unitSel === "shoe" || unitSel === "绿安全国内") return agg.shoeYuan;
  return agg.yuan;
}}

function updateSummary(rows, agg, period, unitSel, custStr, styleStr, chVal, brandVal) {{
  const parts = [];
  parts.push(`${{fmt(rows.length)}} 行`);
  
  const isCompare = period === "sp_compare" || period === "all";
  
  if (isCompare) {{
    parts.push(`${{fmt(agg.custCount)}} 客户`);
    parts.push(`${{fmt(agg.styleCount)}} 鞋型`);
  }} else if (unitSel === "3M") {{
    parts.push(`口罩 ${{fmt(agg.m3Kouzhao)}} 件`);
    parts.push(`耳塞 ${{fmt(agg.m3Ersai)}} 件`);
    parts.push(`其他 ${{fmt(agg.m3Qita)}} 件`);
    parts.push(`${{fmtWan(agg.yuan)}} 万元`);
  }} else if (unitSel === "绿安全国内" && agg.insoleQty > 0) {{
    parts.push(`${{fmt(agg.shoeQty)}} 双`);
    parts.push(`${{fmt(agg.insoleQty)}} 件鞋垫`);
    parts.push(`${{fmtWan(agg.yuan)}} 万元`);
  }} else if (canShowQty(unitSel)) {{
    const displayQty = getDisplayQty(agg, unitSel);
    parts.push(`${{fmt(displayQty)}} ${{qtyUnit(unitSel)}}`);
    parts.push(`${{fmtWan(getDisplayYuan(agg, unitSel))}} 万元`);
  }} else {{
    parts.push(`${{fmtWan(getDisplayYuan(agg, unitSel))}} 万元`);
  }}
  
  if (agg.excRows > 0) {{
    parts.push(`<span class="tag">${{agg.excRows}} 行/${{fmt(agg.excQty)}} 未计营业额</span>`);
  }}
  const summaryHtml = parts.join(" · ");
  document.getElementById("filter-summary").innerHTML = summaryHtml;
  const mobileSummary = document.getElementById("filter-summary-mobile");
  if (mobileSummary) mobileSummary.innerHTML = summaryHtml;
}}

function updateNote(unitSel, agg) {{
  const el = document.getElementById("note-text");
  const box = document.getElementById("note-box");
  let txt = "";
  if (unitSel === "all") {{
    txt = "当前选择「全部单元」，只加营业额。数量（双/件）不可混加。";
    box.classList.add("warn");
  }} else if (unitSel === "3M") {{
    txt = "3M 的口罩、耳塞、其他件数单位不同，不应加总；营业额可以加总。此处只显示营业额合计。";
    box.classList.add("warn");
  }} else {{
    txt = `数量单位：${{qtyUnit(unitSel)}}。金额单位：人民币元（美元已按当月月均汇率折算）。`;
    box.classList.remove("warn");
  }}
  el.textContent = txt;
}}

function updateKPIs(rows, agg, period, unitSel) {{
  const grid = document.getElementById("kpi-grid");
  const showQty = canShowQty(unitSel);
  const unit = qtyUnit(unitSel);
  
  const byY = byYear(rows);
  const cards = [];
  
  if (period === "sp_compare" || period === "all") {{
    const years = [2024, 2025, 2026];
    for (const y of years) {{
      const yRows = byY.get(y) || [];
      const yAgg = aggregate(yRows);
      const prevRows = byY.get(y - 1) || [];
      const prevAgg = aggregate(prevRows);
      const yYuan = getDisplayYuan(yAgg, unitSel);
      const prevYuan = getDisplayYuan(prevAgg, unitSel);
      const yQty = getDisplayQty(yAgg, unitSel);
      const prevQty = getDisplayQty(prevAgg, unitSel);
      
      let label = period === "sp_compare" ? `${{y}} 同期` : `${{y}} 年`;
      if (y === 2026 && period === "all") label = "2026 截至 11-15";
      
      let html = `<div class="card"><div class="k">${{label}}</div>`;
      html += `<div class="v">${{fmtWan(yYuan)}} 万元</div>`;
      if (unitSel === "3M") {{
        html += `<div class="s">口罩 ${{fmt(yAgg.m3Kouzhao)}} / 耳塞 ${{fmt(yAgg.m3Ersai)}} / 其他 ${{fmt(yAgg.m3Qita)}}</div>`;
      }} else if (showQty) {{
        html += `<div class="s">${{fmt(yQty)}} ${{unit}}</div>`;
      }}
      if (y > 2024 && prevYuan) {{
        const yuanCls = pctCls(yYuan, prevYuan);
        html += `<div class="chg ${{yuanCls}}">营业额 ${{pct(yYuan, prevYuan)}}</div>`;
        if (showQty && prevQty) {{
          const qtyCls = pctCls(yQty, prevQty);
          html += `<div class="chg ${{qtyCls}}">数量 ${{pct(yQty, prevQty)}}</div>`;
        }}
      }}
      html += `</div>`;
      cards.push(html);
    }}
  }} else if (unitSel === "3M") {{
    let label = "";
    if (period === "sp2026") label = "2026 同期（1月–11上）";
    else if (period === "sp2025") label = "2025 同期";
    else if (period === "sp2024") label = "2024 同期";
    else if (period === "fy2025") label = "2025 全年";
    else if (period === "fy2024") label = "2024 全年";
    else if (period === "ytd2026") label = "2026 截至 11-15";
    
    cards.push(`<div class="card"><div class="k">${{label}} · 营业额</div><div class="v">${{fmtWan(agg.yuan)}} 万元</div><div class="s">${{fmtYuan(agg.yuan)}} 元</div></div>`);
    cards.push(`<div class="card"><div class="k">${{label}} · 口罩</div><div class="v">${{fmt(agg.m3Kouzhao)}} 件</div></div>`);
    cards.push(`<div class="card"><div class="k">${{label}} · 耳塞</div><div class="v">${{fmt(agg.m3Ersai)}} 件</div></div>`);
    cards.push(`<div class="card"><div class="k">${{label}} · 其他</div><div class="v">${{fmt(agg.m3Qita)}} 件</div></div>`);
    cards.push(`<div class="card"><div class="k">明细行数</div><div class="v">${{fmt(rows.length)}}</div></div>`);
  }} else {{
    let label = "";
    if (period === "sp2026") label = "2026 同期（1月–11上）";
    else if (period === "sp2025") label = "2025 同期";
    else if (period === "sp2024") label = "2024 同期";
    else if (period === "fy2025") label = "2025 全年";
    else if (period === "fy2024") label = "2024 全年";
    else if (period === "ytd2026") label = "2026 截至 11-15";
    
    const displayYuan = getDisplayYuan(agg, unitSel);
    cards.push(`<div class="card"><div class="k">${{label}} · 营业额</div><div class="v">${{fmtWan(displayYuan)}} 万元</div><div class="s">${{fmtYuan(displayYuan)}} 元</div></div>`);
    if (showQty) {{
      cards.push(`<div class="card"><div class="k">${{label}} · 数量</div><div class="v">${{fmt(getDisplayQty(agg, unitSel))}} ${{unit}}</div></div>`);
    }}
    cards.push(`<div class="card"><div class="k">明细行数</div><div class="v">${{fmt(rows.length)}}</div></div>`);
  }}
  
  grid.innerHTML = cards.join("");
}}

function updateSplit(rows, period, unitSel) {{
  const table = document.getElementById("tbl-split");
  const note = document.getElementById("split-note");
  
  let splitKeys = [];
  let splitLabel = "";
  
  if (unitSel === "shoe") {{
    splitKeys = ["内销", "外销", "绿安全华南", "绿安全华东", "绿安全华北", "绿安全外贸", "X"];
    splitLabel = "成品鞋按内销、外销、绿安全国内（华南/华东/华北）、绿安全外贸、X 拆分";
  }} else if (unitSel === "绿安全国内") {{
    splitKeys = ["绿安全华南", "绿安全华东", "绿安全华北"];
    splitLabel = "绿安全国内按华南、华东、华北拆分";
  }} else if (unitSel === "3M") {{
    splitKeys = ["3M口罩", "3M耳塞", "3M其他"];
    splitLabel = "3M 按口罩、耳塞、其他拆分（件数单位不同，不加总）";
  }} else if (unitSel === "all") {{
    splitKeys = FU.filter(f => !f.startsWith("3M") || f === "3M口罩");
    splitKeys = ["内销", "外销", "绿安全华南", "绿安全华东", "绿安全华北", "绿安全外贸", "X",
                 "OEM鞋面/半成品", "钢头", "橡胶底", "3M口罩", "3M耳塞", "3M其他", 
                 "材料", "鞋垫", "其它劳保", "其它未标品牌（材料/配件）"];
    splitLabel = "全部单元按事业/品类拆分（只加营业额）";
  }} else {{
    table.querySelector("thead").innerHTML = "";
    table.querySelector("tbody").innerHTML = "<tr><td>当前筛选无细分项</td></tr>";
    note.textContent = "";
    return;
  }}
  
  note.textContent = splitLabel;
  
  const fuMap = byFu(rows);
  const byY = byYear(rows);
  const years = [...byY.keys()].sort();
  
  const showQty = unitSel !== "all" && unitSel !== "3M";
  
  let thead = "<tr><th>单元</th>";
  for (const y of years) {{ thead += `<th class="num">${{y}} 营业额</th>`; }}
  if (showQty) {{ for (const y of years) {{ thead += `<th class="num">${{y}} 数量</th>`; }} }}
  thead += "</tr>";
  table.querySelector("thead").innerHTML = thead;
  
  let tbody = "";
  const totals = years.map(() => ({{ yuan: 0, qty: 0, shoeQty: 0, shoeYuan: 0 }}));
  
  const useShoeQty = unitSel === "shoe" || unitSel === "绿安全国内";
  
  for (const fu of splitKeys) {{
    const fuRows = fuMap.get(fu) || [];
    if (fuRows.length === 0 && unitSel !== "all") continue;
    
    const fuByY = byYear(fuRows);
    tbody += `<tr><td>${{fu}}</td>`;
    for (let i = 0; i < years.length; i++) {{
      const y = years[i];
      const yRows = fuByY.get(y) || [];
      const yAgg = aggregate(yRows);
      const displayYuan = useShoeQty ? yAgg.shoeYuan : yAgg.yuan;
      tbody += `<td class="num">${{fmtYuan(displayYuan)}}</td>`;
      totals[i].yuan += yAgg.yuan;
      totals[i].shoeYuan += yAgg.shoeYuan;
      totals[i].qty += yAgg.qty;
      totals[i].shoeQty += yAgg.shoeQty;
    }}
    if (showQty) {{
      for (let i = 0; i < years.length; i++) {{
        const y = years[i];
        const yRows = fuByY.get(y) || [];
        const yAgg = aggregate(yRows);
        const displayQ = useShoeQty ? yAgg.shoeQty : yAgg.qty;
        tbody += `<td class="num">${{fmt(displayQ)}}</td>`;
      }}
    }}
    tbody += "</tr>";
  }}
  
  tbody += `<tr class="sum"><td>合计</td>`;
  for (let i = 0; i < years.length; i++) {{
    const displayYuan = useShoeQty ? totals[i].shoeYuan : totals[i].yuan;
    tbody += `<td class="num">${{fmtYuan(displayYuan)}}</td>`;
  }}
  if (showQty) {{
    for (let i = 0; i < years.length; i++) {{
      const displayQ = useShoeQty ? totals[i].shoeQty : totals[i].qty;
      tbody += `<td class="num">${{fmt(displayQ)}}</td>`;
    }}
  }}
  tbody += "</tr>";
  
  table.querySelector("tbody").innerHTML = tbody;
}}

function updateYearChart(rows, unitSel) {{
  const container = document.getElementById("chart-year");
  const byY = byYear(rows);
  const years = [...byY.keys()].sort();
  if (years.length === 0) {{ container.innerHTML = "<p>无数据</p>"; return; }}
  
  const useShoe = unitSel === "shoe" || unitSel === "绿安全国内";
  const data = years.map(y => {{
    const agg = aggregate(byY.get(y) || []);
    return {{ year: y, yuan: useShoe ? agg.shoeYuan : agg.yuan }};
  }});
  const maxVal = Math.max(...data.map(d => d.yuan)) || 1;
  
  const width = 500, height = 200;
  const left = 70, right = 30, top = 20, bottom = 40;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const barW = Math.min(60, plotW / data.length * 0.6);
  const gap = (plotW - barW * data.length) / (data.length + 1);
  
  let svg = `<svg viewBox="0 0 ${{width}} ${{height}}" width="100%">`;
  
  for (let i = 0; i <= 4; i++) {{
    const yy = top + plotH * i / 4;
    const val = maxVal * (1 - i / 4);
    svg += `<line x1="${{left}}" y1="${{yy}}" x2="${{width - right}}" y2="${{yy}}" stroke="#edf0f3"/>`;
    svg += `<text x="${{left - 8}}" y="${{yy + 4}}" text-anchor="end" font-size="10" fill="#7a8490">${{fmtWan(val)}}</text>`;
  }}
  
  data.forEach((d, i) => {{
    const x = left + gap + i * (barW + gap);
    const h = plotH * d.yuan / maxVal;
    const y = top + plotH - h;
    svg += `<rect x="${{x}}" y="${{y}}" width="${{barW}}" height="${{h}}" fill="#1e4d7b" rx="2"/>`;
    svg += `<text x="${{x + barW / 2}}" y="${{height - 10}}" text-anchor="middle" font-size="11" fill="#4a5560">${{d.year}}</text>`;
    svg += `<text x="${{x + barW / 2}}" y="${{y - 4}}" text-anchor="middle" font-size="10" fill="#4a5560">${{fmtWan(d.yuan)}}</text>`;
  }});
  
  svg += `<text x="${{left}}" y="12" font-size="10" fill="#8b95a1">万元</text>`;
  svg += "</svg>";
  container.innerHTML = svg;
}}

function updateMonthChart(rows, unitSel) {{
  const container = document.getElementById("chart-month");
  const byM = byMonth(rows);
  const months = [...byM.keys()].sort();
  if (months.length === 0) {{ container.innerHTML = "<p>无数据</p>"; return; }}
  
  const useShoe = unitSel === "shoe" || unitSel === "绿安全国内";
  const data = months.map(ym => {{
    const agg = aggregate(byM.get(ym) || []);
    return {{ ym, yuan: useShoe ? agg.shoeYuan : agg.yuan }};
  }});
  const maxVal = Math.max(...data.map(d => d.yuan)) || 1;
  
  const width = 800, height = 220;
  const left = 60, right = 20, top = 20, bottom = 50;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  
  let svg = `<svg viewBox="0 0 ${{width}} ${{height}}" width="100%">`;
  
  for (let i = 0; i <= 4; i++) {{
    const yy = top + plotH * i / 4;
    const val = maxVal * (1 - i / 4);
    svg += `<line x1="${{left}}" y1="${{yy}}" x2="${{width - right}}" y2="${{yy}}" stroke="#edf0f3"/>`;
    svg += `<text x="${{left - 8}}" y="${{yy + 4}}" text-anchor="end" font-size="10" fill="#7a8490">${{fmtWan(val)}}</text>`;
  }}
  
  const pts = data.map((d, i) => {{
    const x = left + (plotW * i / (data.length - 1 || 1));
    const y = top + plotH - plotH * d.yuan / maxVal;
    return {{ x, y, ym: d.ym }};
  }});
  
  svg += `<polyline fill="none" stroke="#1e4d7b" stroke-width="2" points="${{pts.map(p => p.x + "," + p.y).join(" ")}}"/>`;
  pts.forEach(p => {{
    const fill = p.ym === "2026-11" ? "#b42318" : "#1e4d7b";
    svg += `<circle cx="${{p.x}}" cy="${{p.y}}" r="3" fill="${{fill}}"/>`;
  }});
  
  const labelStep = Math.ceil(data.length / 12);
  data.forEach((d, i) => {{
    if (i % labelStep === 0 || i === data.length - 1) {{
      const x = left + (plotW * i / (data.length - 1 || 1));
      svg += `<text x="${{x}}" y="${{height - 10}}" text-anchor="middle" font-size="9" fill="#4a5560">${{d.ym.slice(2)}}</text>`;
    }}
  }});
  
  svg += `<text x="${{left}}" y="12" font-size="10" fill="#8b95a1">万元</text>`;
  svg += "</svg>";
  container.innerHTML = svg;
}}

function updateMonthTable(rows, unitSel) {{
  const table = document.getElementById("tbl-month");
  const byM = byMonth(rows);
  const months = [...byM.keys()].sort();
  
  const showQty = canShowQty(unitSel);
  const unit = qtyUnit(unitSel);
  
  let thead = "<tr><th>月份</th><th class=\\"num\\">营业额（元）</th>";
  if (showQty) thead += `<th class="num">数量（${{unit}}）</th>`;
  thead += "<th class=\\"num\\">行数</th></tr>";
  table.querySelector("thead").innerHTML = thead;
  
  const useShoe = unitSel === "shoe" || unitSel === "绿安全国内";
  let tbody = "";
  let totalYuan = 0, totalQty = 0, totalRows = 0;
  for (const ym of months) {{
    const mRows = byM.get(ym) || [];
    const mAgg = aggregate(mRows);
    const displayYuan = useShoe ? mAgg.shoeYuan : mAgg.yuan;
    totalYuan += displayYuan;
    const displayQ = useShoe ? mAgg.shoeQty : mAgg.qty;
    totalQty += displayQ;
    totalRows += mRows.length;
    
    tbody += `<tr><td>${{ym}}</td><td class="num">${{fmtYuan(displayYuan)}}</td>`;
    if (showQty) tbody += `<td class="num">${{fmt(displayQ)}}</td>`;
    tbody += `<td class="num">${{fmt(mRows.length)}}</td></tr>`;
  }}
  tbody += `<tr class="sum"><td>合计</td><td class="num">${{fmtYuan(totalYuan)}}</td>`;
  if (showQty) tbody += `<td class="num">${{fmt(totalQty)}}</td>`;
  tbody += `<td class="num">${{fmt(totalRows)}}</td></tr>`;
  
  table.querySelector("tbody").innerHTML = tbody;
}}

function updateDimTable(tableId, rows, dimIdx, dimList, limit, unitSel) {{
  const table = document.getElementById(tableId);
  const byD = byDim(rows, dimIdx);
  
  const showQty = canShowQty(unitSel);
  const unit = qtyUnit(unitSel);
  
  const useShoe = unitSel === "shoe" || unitSel === "绿安全国内";
  const items = [...byD.entries()].map(([idx, dRows]) => {{
    const agg = aggregate(dRows);
    const displayYuan = useShoe ? agg.shoeYuan : agg.yuan;
    const displayQ = useShoe ? agg.shoeQty : agg.qty;
    return {{ name: dimList[idx], yuan: displayYuan, qty: displayQ, rows: dRows.length }};
  }}).sort((a, b) => b.yuan - a.yuan).slice(0, limit);
  
  let thead = "<tr><th>名称</th><th class=\\"num\\">营业额（元）</th>";
  if (showQty) thead += `<th class="num">数量（${{unit}}）</th>`;
  thead += "<th class=\\"num\\">行数</th></tr>";
  table.querySelector("thead").innerHTML = thead;
  
  let tbody = "";
  for (const item of items) {{
    const tag = (tableId === "tbl-cust" && item.name === "NEWTEX1") ? '<span class="tag">10月后口径异常</span>' : "";
    tbody += `<tr><td>${{item.name}}${{tag}}</td><td class="num">${{fmtYuan(item.yuan)}}</td>`;
    if (showQty) tbody += `<td class="num">${{fmt(item.qty)}}</td>`;
    tbody += `<td class="num">${{fmt(item.rows)}}</td></tr>`;
  }}
  
  table.querySelector("tbody").innerHTML = tbody;
}}

function initDataLists() {{
  const dlCust = document.getElementById("dl-cust");
  dlCust.innerHTML = CUST.map(c => `<option value="${{c}}"/>`).join("");
  
  const dlStyle = document.getElementById("dl-style");
  dlStyle.innerHTML = STYLE.map(s => `<option value="${{s}}"/>`).join("");
}}

function initFxTable() {{
  const el = document.getElementById("fx-table");
  const rows = FX.map(f => `${{f.ym}}: ${{f.rate.toFixed(4)}}`);
  el.innerHTML = rows.join(" · ");
}}

function resetFilters() {{
  document.getElementById("f-period").value = "sp2026";
  document.getElementById("f-unit").value = "shoe";
  document.getElementById("f-cust").value = "";
  document.getElementById("f-style").value = "";
  document.getElementById("f-ch").value = "";
  document.getElementById("f-brand").value = "";
  render();
}}

function debounce(fn, ms) {{
  let t;
  return function(...args) {{
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), ms);
  }};
}}

function toggleFiltersExtra() {{
  const extra = document.getElementById("filters-extra");
  const btn = document.getElementById("btn-expand");
  if (extra.classList.contains("show")) {{
    extra.classList.remove("show");
    btn.textContent = "更多筛选 ▼";
  }} else {{
    extra.classList.add("show");
    btn.textContent = "收起筛选 ▲";
  }}
}}

function updateMobileSummary() {{
  const el = document.getElementById("filter-summary-mobile");
  const summary = document.getElementById("filter-summary").innerHTML;
  el.innerHTML = summary;
}}

document.addEventListener("DOMContentLoaded", function() {{
  initDataLists();
  initFxTable();
  
  document.getElementById("f-period").addEventListener("change", render);
  document.getElementById("f-unit").addEventListener("change", render);
  document.getElementById("f-cust").addEventListener("input", debounce(render, 300));
  document.getElementById("f-style").addEventListener("input", debounce(render, 300));
  document.getElementById("f-ch").addEventListener("change", render);
  document.getElementById("f-brand").addEventListener("change", render);
  document.getElementById("btn-reset").addEventListener("click", resetFilters);
  document.getElementById("btn-expand").addEventListener("click", toggleFiltersExtra);
  
  document.getElementById("loading").classList.add("hidden");
  document.getElementById("main-content").classList.remove("hidden");
  
  render();
}});
</script>
</body>
</html>'''
    
    return html


def main():
    print("Loading FX rates...")
    rates = load_fx_rates()
    print(f"  Loaded {len(rates)} monthly rates")
    
    print("Loading and merging Excel data...")
    df = load_and_merge_data()
    print(f"  Merged {len(df)} rows")
    
    print("Applying FX rates and classifying rows...")
    df = apply_fx_and_classify(df, rates)
    
    unassigned = (df["unit"] == "未分类").sum()
    if unassigned > 0:
        print(f"  WARNING: {unassigned} unassigned rows")
        print(df[df["unit"] == "未分类"][["日期", "类别", "品牌", "客户简称", "鞋型"]].head(10))
    
    total_rows = len(df)
    total_qty = df["qty"].sum()
    print(f"  Total rows: {total_rows}, Total qty: {total_qty}")
    
    sp2026_shoe = df[(df["year"] == 2026) & df["same"] & df["is_shoe"]]
    sp2026_shoe_agg = sp2026_shoe[sp2026_shoe["amt_flag"] == ""]
    qty = sp2026_shoe["qty"].sum()
    fen = sp2026_shoe_agg["amt_fen"].sum()
    yuan = fen / 100
    print(f"  2026 同期 成品鞋: {qty} 双, {yuan:.2f} 元 ({int(round(yuan))} 元)")
    
    expected_qty = 719370
    expected_yuan = 70409073
    if qty == expected_qty and abs(yuan - expected_yuan) < 1:
        print(f"  ✓ Validation PASSED: qty={qty}, yuan={int(round(yuan))}")
    else:
        print(f"  ✗ Validation FAILED: expected {expected_qty} 双 / {expected_yuan} 元")
        print(f"    got {qty} 双 / {int(round(yuan))} 元")
    
    print("Packing data for JavaScript...")
    data = pack_data_for_js(df, rates)
    print(f"  Rows for JS: {len(data['rows'])}")
    
    print("Generating HTML...")
    html = generate_html(data)
    
    out_path = OUT / "zr-full-sales-v4.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"  Written to: {out_path}")
    print(f"  File size: {out_path.stat().st_size:,} bytes")
    
    return df, data


if __name__ == "__main__":
    main()
