# /// script
# requires-python = ">=3.10"
# dependencies = ["pandas", "xlrd"]
# ///
"""把戶政司「縣市出生死亡結婚離婚(按登記)」民國 100–115 年的月報 xls，轉成整齊的月資料 CSV，並做加總核對。

用法：uv run etl/etl.py
輸出：data/monthly.csv、data/region_mapping.csv、data/validation_report.txt
"""
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "raw"
OUT = ROOT / "data"

# 縣市對照：標準名稱 -> (區域, 類型)。區域分法是人工判斷。
REGIONS = {
    "新北市": ("北部", "直轄市"), "臺北市": ("北部", "直轄市"), "桃園市": ("北部", "直轄市"),
    "基隆市": ("北部", "市"), "新竹市": ("北部", "市"), "新竹縣": ("北部", "縣"), "宜蘭縣": ("北部", "縣"),
    "臺中市": ("中部", "直轄市"), "苗栗縣": ("中部", "縣"), "彰化縣": ("中部", "縣"),
    "南投縣": ("中部", "縣"), "雲林縣": ("中部", "縣"),
    "臺南市": ("南部", "直轄市"), "高雄市": ("南部", "直轄市"), "嘉義市": ("南部", "市"),
    "嘉義縣": ("南部", "縣"), "屏東縣": ("南部", "縣"),
    "花蓮縣": ("東部", "縣"), "臺東縣": ("東部", "縣"),
    "澎湖縣": ("離島", "縣"), "金門縣": ("離島", "縣"), "連江縣": ("離島", "縣"),
}
ALIAS = {"桃園縣": "桃園市"}  # 103/12/25 改制為直轄市，視為同一個縣市
FUJIAN = ["金門縣", "連江縣"]
TAIWAN_PROVINCE = [r for r, (_, t) in REGIONS.items() if t != "直轄市" and r not in FUJIAN]
AGG = {"總計", "臺灣省", "福建省"}

clean = lambda s: re.sub(r"[\s　]", "", str(s))


def num(v, where):
    if pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().replace(",", "")
    if s in ("-", "－", "—"):
        return 0
    try:
        return int(float(s))
    except ValueError:
        raise ValueError(f"{where}: 無法解析數字 {v!r}")


def find_cols(d, where):
    """從表頭（第 3 列）依文字找出各欄位置，不同年份的欄位數不同。"""
    head = {i: clean(v) for i, v in d.iloc[2].items() if pd.notna(v)}

    def pick(prefix):
        hit = next((i for i, t in head.items() if t.startswith(prefix)), None)
        if hit is None:
            raise ValueError(f"{where}: 表頭找不到「{prefix}」")
        return hit

    return {
        "births_total": pick("出生總數"), "deaths": pick("死亡數"),
        "marriages": pick("結婚"), "divorces": pick("離婚"),
        # 出生總數底下的分項：婚生、非婚生已認領、非婚生未認領（各年固定在 3、4、5 欄）
        "births_legit": 3, "births_ack": 4, "births_unack": 5,
    }


def triples(d, start=3):
    """每個區域佔 3 列（計、男、女），名稱只寫在其中一列（合併儲存格）。"""
    sex = d[1].map(lambda v: clean(v) if pd.notna(v) else None)
    idx = [i for i in d.index[start:] if sex[i] in ("計", "男", "女")]
    if len(idx) % 3:
        raise ValueError(f"性別列數 {len(idx)} 不是 3 的倍數")
    for k in range(0, len(idx), 3):
        t = idx[k:k + 3]
        if [sex[i] for i in t] != ["計", "男", "女"]:
            raise ValueError(f"第 {t[0]} 列起的性別順序不是 計/男/女")
        label = next((clean(d.at[i, 0]) for i in t if pd.notna(d.at[i, 0])), None)
        if label is None:
            raise ValueError(f"第 {t[0]} 列起找不到區域名稱")
        yield label, t


def parse_sheet(path, sheet):
    where = f"{path.name}/{sheet}"
    d = pd.read_excel(path, sheet_name=sheet, header=None)
    m = re.search(r"(\d+)年\s*(\d+)月", " ".join(str(v) for v in d.iloc[1].dropna()))
    if not m:
        return None  # 不是月表（累計、比率表等）
    year, month = int(m.group(1)), int(m.group(2))
    cols = find_cols(d, where)
    recs = []
    for label, (t, ma, fe) in triples(d):
        g = lambda i, key: num(d.at[i, cols[key]], f"{where} 列{i} {key}")
        gx = lambda key: next((v for v in (g(i, key) for i in (t, ma, fe)) if v is not None), None)
        recs.append(dict(
            year_roc=year, month=month, region_raw=label, region=ALIAS.get(label, label),
            births_total=g(t, "births_total"), births_male=g(ma, "births_total"), births_female=g(fe, "births_total"),
            births_legit=g(t, "births_legit"),
            births_nonlegit=(g(t, "births_ack") or 0) + (g(t, "births_unack") or 0),
            deaths_total=g(t, "deaths"), deaths_male=g(ma, "deaths"), deaths_female=g(fe, "deaths"),
            marriages=gx("marriages"), divorces=gx("divorces"),
        ))
    return recs


def parse_cumulative(path, sheet):
    """全年累計表（用來核對 12 個月加總）：{區域: (出生, 死亡, 結婚, 離婚)}。"""
    d = pd.read_excel(path, sheet_name=sheet, header=None)
    cols = find_cols(d, f"{path.name}/{sheet}")
    out = {}
    for label, trio in triples(d):
        pick = lambda c: next((v for v in (num(d.at[i, cols[c]], "累計") for i in trio) if v is not None), None)
        out[ALIAS.get(label, label)] = tuple(pick(c) for c in ("births_total", "deaths", "marriages", "divorces"))
    return out


def parse_pop_sheet(path, sheet):
    """月底現住人口（按性別及單一年齡）：只取「總計」欄，回傳每個區域的 (計, 男, 女) 人口。"""
    where = f"{path.name}/{sheet}"
    d = pd.read_excel(path, sheet_name=sheet, header=None)
    m = re.search(r"(\d+)年\s*(\d+)月", " ".join(str(v) for v in d.iloc[1].dropna()))
    if not m:
        raise ValueError(f"{where}: 找不到月份標題")
    col = next((i for i, v in d.iloc[2].items() if pd.notna(v) and clean(v) == "總計"), None)
    if col is None:
        raise ValueError(f"{where}: 表頭找不到「總計」欄")
    recs = []
    for label, (t, ma, fe) in triples(d):
        recs.append(dict(year_roc=int(m.group(1)), month=int(m.group(2)), region_raw=label, region=ALIAS.get(label, label),
                         pop_total=num(d.at[t, col], where), pop_male=num(d.at[ma, col], where), pop_female=num(d.at[fe, col], where)))
    return recs


def parse_official_rates(path, sheet):
    """報表附的「全年各縣市出生、死亡人數及比率」：{區域: (出生數, 粗出生率, 死亡數, 粗死亡率)}。"""
    d = pd.read_excel(path, sheet_name=sheet, header=None)
    out = {}
    for i in d.index[4:]:
        label = clean(d.at[i, 0]) if pd.notna(d.at[i, 0]) else None
        if label and pd.notna(d.at[i, 1]) and not label.startswith("說明"):
            out[ALIAS.get(label, label)] = tuple(float(d.at[i, c]) for c in (1, 2, 3, 4))
    return out


def main():
    rows, agg_rows, cum, log = [], [], {}, []
    official = {}
    files = sorted(RAW.glob("縣市出生死亡結婚離婚*.xls"), key=lambda p: int(re.search(r"-(\d+)年", p.name).group(1)))
    for f in files:
        fy = int(re.search(r"-(\d+)年", f.name).group(1))
        months = []
        for sh in pd.ExcelFile(f).sheet_names:
            if re.search(r"累計|合計", sh):  # 全年累計表，只用來核對
                cum[fy] = parse_cumulative(f, sh)
                continue
            if "出生率" in sh:  # 官方全年比率表，只用來核對人口數算出的比率
                official[fy] = parse_official_rates(f, sh)
                continue
            recs = parse_sheet(f, sh)
            if recs is None:
                continue
            if recs[0]["year_roc"] != fy:
                raise ValueError(f"{f.name}/{sh}: 表內年份 {recs[0]['year_roc']} 和檔名 {fy} 不同")
            months.append(recs[0]["month"])
            for r in recs:
                (rows if r["region"] in REGIONS else agg_rows).append(r)
        if months != list(range(1, len(months) + 1)):
            raise ValueError(f"{f.name}: 月份不連續 {months}")
        log.append(f"民國{fy}年：{len(months)} 個月（1–{months[-1]} 月）")

    # ---------- 人口 ----------
    pop_rows, pop_agg = [], []
    for f in sorted(RAW.glob("縣市人口數按性別及年齡*.xls"), key=lambda p: int(re.search(r"-(\d+)年", p.name).group(1))):
        for sh in pd.ExcelFile(f).sheet_names:
            for r in parse_pop_sheet(f, sh):
                (pop_rows if r["region"] in REGIONS else pop_agg).append(r)
    pop, pop_agg = pd.DataFrame(pop_rows), pd.DataFrame(pop_agg)

    df, agg = pd.DataFrame(rows), pd.DataFrame(agg_rows)
    n0 = len(df)
    df = df.merge(pop[["year_roc", "month", "region", "pop_total", "pop_male", "pop_female"]],
                  on=["year_roc", "month", "region"], how="left", validate="one_to_one")
    assert len(df) == n0 and df.pop_total.notna().all(), "有月份或縣市找不到人口數"
    unknown = set(agg.region) - AGG
    if unknown:
        raise ValueError(f"出現未知的區域名稱：{unknown}")
    df["year_ad"] = df.year_roc + 1911
    df["area"] = df.region.map(lambda r: REGIONS[r][0])
    cols = ["year_roc", "year_ad", "month", "region", "area", "region_raw", "births_total", "births_male",
            "births_female", "births_legit", "births_nonlegit", "deaths_total", "deaths_male", "deaths_female",
            "marriages", "divorces", "pop_total", "pop_male", "pop_female"]
    df = df[cols].sort_values(["year_roc", "month", "region"]).reset_index(drop=True)

    # ---------- 驗證 ----------
    nmonths = df[["year_roc", "month"]].drop_duplicates().shape[0]
    rep = ["== ETL 驗證報告 ==", *log,
           f"月資料 {len(df)} 列 = {df.region.nunique()} 個縣市 × {nmonths} 個月", ""]
    bad = 0
    key = ["year_roc", "month"]
    vals = ["births_total", "deaths_total", "marriages", "divorces"]

    for kind in ("births", "deaths"):  # 1. 男 + 女 = 合計
        diff = df[df[f"{kind}_male"] + df[f"{kind}_female"] != df[f"{kind}_total"]]
        rep.append(f"[{'OK' if diff.empty else '不符'}] {kind}：男 + 女 = 合計（不符 {len(diff)} 列）")
        bad += len(diff)

    tot = agg[agg.region == "總計"].set_index(key)[vals]  # 2. 22 縣市加總 = 全國總計
    s = df.groupby(key)[vals].sum()
    d = (s - tot).abs()
    for v in vals:
        wrong = d[(d[v] != 0) | d[v].isna()][v]
        rep.append(f"[{'OK' if wrong.empty else '不符'}] 22 縣市加總 = 全國總計：{v}（不符 {len(wrong)} 個月）")
        for (y, m), x in list(wrong.items())[:10]:
            rep.append(f"      民國{y}年{m}月 差 {x}（縣市加總 {s.at[(y, m), v]}，報表總計 {tot.at[(y, m), v]}）")
        bad += len(wrong)

    for name, members in (("臺灣省", TAIWAN_PROVINCE), ("福建省", FUJIAN)):  # 3. 省 = 轄下縣市加總
        a = agg[agg.region == name].set_index(key)[vals]
        sel = df.region.isin(members)
        if name == "臺灣省":  # 桃園縣在 103/12 升格前仍屬臺灣省
            sel |= (df.region == "桃園市") & ((df.year_roc < 103) | ((df.year_roc == 103) & (df.month < 12)))
        sm = df[sel].groupby(key)[vals].sum()
        n = int(((sm - a).abs() != 0).any(axis=1).sum())
        rep.append(f"[{'OK' if n == 0 else '不符'}] {name} = 轄下縣市加總（不符 {n} 個月）")
        bad += n

    for fy, c in sorted(cum.items()):  # 4. 各月加總 = 全年累計表
        sub = df[df.year_roc == fy].groupby("region")[vals].sum()
        nbad = sum(1 for r, t in c.items() if r in sub.index and tuple(int(x) for x in sub.loc[r]) != tuple(t))
        rep.append(f"[{'OK' if nbad == 0 else '不符'}] 民國{fy}年：各月加總 = 全年累計表（不符 {nbad} 個縣市）")
        bad += nbad

    # 5. 人口：男 + 女 = 計；22 縣市加總 = 全國總計；臺灣省、福建省
    dpop = df[df.pop_male + df.pop_female != df.pop_total]
    rep.append(f"[{'OK' if dpop.empty else '不符'}] 人口：男 + 女 = 合計（不符 {len(dpop)} 列）")
    bad += len(dpop)
    ptot = pop_agg[pop_agg.region == "總計"].set_index(key)["pop_total"]
    ps = df.groupby(key)["pop_total"].sum()
    n = int(((ps - ptot).abs() != 0).sum())
    rep.append(f"[{'OK' if n == 0 else '不符'}] 人口：22 縣市加總 = 全國總計（不符 {n} 個月）")
    bad += n
    for name, members in (("臺灣省", TAIWAN_PROVINCE), ("福建省", FUJIAN)):
        a = pop_agg[pop_agg.region == name].set_index(key)["pop_total"]
        sel = df.region.isin(members)
        if name == "臺灣省":
            sel |= (df.region == "桃園市") & ((df.year_roc < 103) | ((df.year_roc == 103) & (df.month < 12)))
        n = int(((df[sel].groupby(key)["pop_total"].sum() - a).abs() != 0).sum())
        rep.append(f"[{'OK' if n == 0 else '不符'}] 人口：{name} = 轄下縣市加總（不符 {n} 個月）")
        bad += n
    # 6. 用人口數算出的全年粗出生率、粗死亡率，對照報表附的官方比率（民國 107–113 年）。
    #    官方分母是年中人口，這裡用 12 個月底人口的平均，所以允許小誤差。
    worst = (0.0, "")
    for fy, off in sorted(official.items()):
        yr = df[df.year_roc == fy].groupby("region").agg(b=("births_total", "sum"), d=("deaths_total", "sum"), p=("pop_total", "mean"))
        for r, (ob, ocbr, od, ocdr) in off.items():
            if r not in yr.index:
                continue
            if int(yr.at[r, "b"]) != int(ob) or int(yr.at[r, "d"]) != int(od):
                rep.append(f"      民國{fy}年 {r}：出生/死亡數與官方表不同（{yr.at[r, 'b']}/{yr.at[r, 'd']} vs {int(ob)}/{int(od)}）"); bad += 1
            for mine, theirs in ((yr.at[r, "b"] / yr.at[r, "p"] * 1000, ocbr), (yr.at[r, "d"] / yr.at[r, "p"] * 1000, ocdr)):
                if abs(mine - theirs) > worst[0]:
                    worst = (abs(mine - theirs), f"民國{fy}年 {r}：我算 {mine:.3f}‰，官方 {theirs:.3f}‰")
    ok = worst[0] <= 0.1
    rep.append(f"[{'OK' if ok else '不符'}] 全年粗出生率、粗死亡率 vs 官方比率表（民國 {min(official)}–{max(official)} 年）：最大差 {worst[0]:.3f}‰（{worst[1]}）")
    bad += 0 if ok else 1
    rep += ["", "全部通過" if bad == 0 else f"共 {bad} 項不符，詳見上方"]

    OUT.mkdir(exist_ok=True)
    df.to_csv(OUT / "monthly.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame([(r, a, t) for r, (a, t) in REGIONS.items()], columns=["region", "area", "type"]) \
        .to_csv(OUT / "region_mapping.csv", index=False, encoding="utf-8-sig")
    (OUT / "validation_report.txt").write_text("\n".join(rep) + "\n", encoding="utf-8")
    print("\n".join(rep))
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
