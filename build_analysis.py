"""
Auspify Business Analytics Internship - Netflix dataset
Builds Tasks 1-4 (Excel workbooks with live formulas + native charts), PNG charts and a cleaned CSV.
Run:  python build_analysis.py
"""
import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList

SRC = "/mnt/user-data/uploads/Dataset.csv"
OUT = "/mnt/user-data/outputs/Auspify_Business_Analytics"
PNG = os.path.join(OUT, "charts")
os.makedirs(PNG, exist_ok=True)

# ----------------------------------------------------------------- 1. CLEAN
SEGMENT = {"TV-Y": "Kids", "TV-Y7": "Kids", "TV-Y7-FV": "Kids", "TV-G": "Kids", "G": "Kids",
           "TV-PG": "Family (Parental Guidance)", "PG": "Family (Parental Guidance)",
           "TV-14": "Teens", "PG-13": "Teens",
           "TV-MA": "Adults", "R": "Adults", "NC-17": "Adults",
           "NR": "Unrated", "UR": "Unrated"}
SEG_ORDER = ["Kids", "Family (Parental Guidance)", "Teens", "Adults", "Unrated"]
REGION = {}
for reg, lst in {
    "North America": ["United States", "Canada"],
    "Latin America": ["Argentina", "Brazil", "Chile", "Colombia", "Guatemala", "Jamaica", "Mexico", "Paraguay",
                      "Peru", "Puerto Rico", "Uruguay", "Venezuela"],
    "Europe": ["Austria", "Belarus", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czech Republic", "Denmark",
               "Finland", "France", "Georgia", "Germany", "Greece", "Hungary", "Iceland", "Ireland", "Italy",
               "Luxembourg", "Netherlands", "Norway", "Poland", "Portugal", "Romania", "Russia", "Serbia",
               "Slovenia", "Soviet Union", "Spain", "Sweden", "Switzerland", "Ukraine", "United Kingdom",
               "West Germany"],
    "South Asia": ["India", "Pakistan", "Bangladesh"],
    "East & Southeast Asia": ["Cambodia", "China", "Hong Kong", "Indonesia", "Japan", "Malaysia", "Philippines",
                              "Singapore", "South Korea", "Taiwan", "Thailand", "Vietnam"],
    "Oceania": ["Australia", "New Zealand"],
    "Middle East & North Africa": ["Egypt", "Iran", "Israel", "Jordan", "Kuwait", "Lebanon", "Saudi Arabia",
                                   "Syria", "Turkey", "United Arab Emirates"],
    "Sub-Saharan Africa": ["Cameroon", "Ghana", "Kenya", "Mauritius", "Mozambique", "Namibia", "Nigeria",
                           "Senegal", "Somalia", "South Africa", "Zimbabwe"],
}.items():
    for c in lst:
        REGION[c] = reg

df = pd.read_csv(SRC).drop_duplicates()
df["date_added"] = pd.to_datetime(df["date_added"], format="mixed")
df["year_added"] = df["date_added"].dt.year
df["audience_segment"] = df["rating"].map(SEGMENT)
df["duration_value"] = df["duration"].str.extract(r"(\d+)").astype(int)
df["region"] = df["country"].map(REGION).fillna("Unknown")
COLS = ["show_id", "type", "title", "director", "country", "date_added", "year_added", "release_year", "rating",
        "audience_segment", "duration", "duration_value", "listed_in", "region"]
df = df[COLS]
df.to_csv(os.path.join(OUT, "Dataset_Cleaned.csv"), index=False)
N = len(df)
LET = {c: chr(65 + i) for i, c in enumerate(COLS)}


def R(col):
    return f"Data!${LET[col]}$2:${LET[col]}${N + 1}"


MOV, TV = "Movie", "TV Show"
tot_m, tot_t = (df.type == MOV).sum(), (df.type == TV).sum()
known = df[df.country != "Not Given"]

# ----------------------------------------------------------------- 2. STYLE HELPERS
F = "Arial"
RED, DARK, GREY = "B20710", "221F1F", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(**k):
    return Font(name=F, size=k.pop("size", 10), **k)


def new_wb():
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = font(size=16, bold=True, color=RED)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = font(italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def table(ws, r0, c0, headers, rows, fmts=None, total_row=False):
    for j, h in enumerate(headers):
        c = ws.cell(r0, c0 + j, h)
        c.font = font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=RED)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = ws.cell(r0 + 1 + i, c0 + j, v)
            c.font = font(bold=(total_row and i == len(rows) - 1))
            c.border = BORDER
            if total_row and i == len(rows) - 1:
                c.fill = PatternFill("solid", fgColor=GREY)
            if fmts and fmts[j]:
                c.number_format = fmts[j]
    return r0 + len(rows)


def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = font(italic=True, size=9, color="595959")


def data_sheet(wb):
    ws = wb.create_sheet("Data")
    for j, c in enumerate(COLS, 1):
        h = ws.cell(1, j, c)
        h.font = font(bold=True, color="FFFFFF")
        h.fill = PatternFill("solid", fgColor=DARK)
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v.to_pydatetime() if isinstance(v, pd.Timestamp) else v)
            if j == 6:
                c.number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{chr(64 + len(COLS))}{N + 1}"
    for j, w in enumerate([9, 9, 34, 24, 18, 12, 11, 12, 9, 24, 11, 14, 40, 26], 1):
        ws.column_dimensions[chr(64 + j)].width = w


def insights(wb, heading, lines):
    ws = wb.create_sheet("Insights", 0 if False else None)
    title(ws, heading, "Observations are generated from the dataset; figures in the tables update by formula, this text does not.")
    r = 4
    for kind, text in lines:
        c = ws.cell(r, 1, text)
        if kind == "h":
            c.font = font(bold=True, size=12, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=RED)
        else:
            c.font = font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * (1 + len(text) // 120)
        r += 1
    ws.column_dimensions["A"].width = 130


def bar(ws, ttl, cats, data, anchor, stacked=False, horizontal=False, w=18, h=9, ytitle=None, fmt=None):
    c = BarChart()
    c.type = "bar" if horizontal else "col"
    c.title = ttl
    if stacked:
        c.grouping, c.overlap = "stacked", 100
    c.add_data(data, titles_from_data=True)
    c.set_categories(cats)
    c.width, c.height = w, h
    c.x_axis.delete = c.y_axis.delete = False
    if horizontal:
        c.x_axis.scaling.orientation = "maxMin"
    if ytitle:
        c.y_axis.title = ytitle
    if fmt:
        c.y_axis.number_format = fmt
    ws.add_chart(c, anchor)


def line(ws, ttl, cats, data, anchor, w=18, h=9, fmt=None):
    c = LineChart()
    c.title = ttl
    c.add_data(data, titles_from_data=True)
    c.set_categories(cats)
    c.width, c.height = w, h
    c.x_axis.delete = c.y_axis.delete = False
    if fmt:
        c.y_axis.number_format = fmt
    ws.add_chart(c, anchor)


def pie(ws, ttl, cats, data, anchor, w=12, h=9):
    c = PieChart()
    c.title = ttl
    c.add_data(data, titles_from_data=True)
    c.set_categories(cats)
    c.dataLabels = DataLabelList()
    c.dataLabels.showPercent = True
    c.width, c.height = w, h
    ws.add_chart(c, anchor)


def png(fig, name):
    fig.tight_layout()
    fig.savefig(os.path.join(PNG, name), dpi=150)
    plt.close(fig)


def pct(x):
    return f"{x:.1%}"


TYPE_COL = {MOV: "#B20710", TV: "#221F1F"}

# ================================================================= TASK 1
def task1():
    wb = new_wb()
    cnt = lambda col, v: f'COUNTIF({R(col)},"{v}")'
    # ---- KPI Summary
    ws = wb.create_sheet("KPI Summary")
    title(ws, "Task 1 - Netflix Content Portfolio Analysis", "KPI summary (all values are live formulas on the Data sheet)")
    k = [
        ("Total titles in catalogue", f"=COUNTA({R('show_id')})", "#,##0", "Rows in cleaned dataset"),
        ("Movies", f'=COUNTIF({R("type")},"Movie")', "#,##0", ""),
        ("TV Shows", f'=COUNTIF({R("type")},"TV Show")', "#,##0", ""),
        ("Movie share of catalogue", "=B5/B4", "0.0%", ""),
        ("TV Show share of catalogue", "=B6/B4", "0.0%", ""),
        ("Movies per TV Show (ratio)", "=B5/B6", "0.00", "Content-mix ratio"),
        ("Average movie runtime (minutes)", f'=AVERAGEIFS({R("duration_value")},{R("type")},"Movie")', "0.0", "From the duration column"),
        ("Average seasons per TV Show", f'=AVERAGEIFS({R("duration_value")},{R("type")},"TV Show")', "0.00", ""),
        ("TV Shows with only 1 season", f'=COUNTIFS({R("type")},"TV Show",{R("duration_value")},1)/B6', "0.0%", "Share of all TV titles"),
        ("Earliest release year", f"=MIN({R('release_year')})", "0", ""),
        ("Latest release year", f"=MAX({R('release_year')})", "0", ""),
        ("Titles added to Netflix 2019-2021", f'=COUNTIFS({R("year_added")},">=2019")/B4', "0.0%", "Share of catalogue (2021 is partial, data ends 25-Sep-2021)"),
        ("Countries represented", int(known.country.nunique()), "0", "Hard-coded count (distinct values excl. 'Not Given')"),
        ("Distinct genre tags", int(df.listed_in.str.split(", ").explode().nunique()), "0", "Hard-coded count after splitting listed_in"),
    ]
    table(ws, 3, 1, ["KPI", "Value", "Note"], [(a, b, d) for a, b, c, d in k])
    for i, (a, b, fm, d) in enumerate(k):
        ws.cell(4 + i, 2).number_format = fm
    widths(ws, {"A": 38, "B": 16, "C": 62})

    # ---- Content Mix
    ws = wb.create_sheet("Content Mix")
    title(ws, "Content Mix: Movies vs TV Shows")
    table(ws, 3, 1, ["Content Type", "Titles", "% Share"],
          [("Movie", f'=COUNTIF({R("type")},A4)', "=B4/$B$6"),
           ("TV Show", f'=COUNTIF({R("type")},A5)', "=B5/$B$6"),
           ("Total", "=SUM(B4:B5)", "=SUM(C4:C5)")],
          [None, "#,##0", "0.0%"], total_row=True)
    eras = [("Before 2000", 1900, 1999), ("2000-2009", 2000, 2009), ("2010-2014", 2010, 2014),
            ("2015-2017", 2015, 2017), ("2018-2021", 2018, 2021)]
    rows = []
    for i, (e, lo, hi) in enumerate(eras):
        r = 11 + i
        rows.append((e, lo, hi,
                     f'=COUNTIFS({R("type")},"Movie",{R("release_year")},">="&B{r},{R("release_year")},"<="&C{r})',
                     f'=COUNTIFS({R("type")},"TV Show",{R("release_year")},">="&B{r},{R("release_year")},"<="&C{r})',
                     f"=D{r}+E{r}", f"=E{r}/F{r}"))
    ws["A9"] = "Catalogue by release era"
    ws["A9"].font = font(bold=True, size=12)
    table(ws, 10, 1, ["Release Era", "From", "To", "Movies", "TV Shows", "Total", "TV Show %"], rows,
          [None, "0", "0", "#,##0", "#,##0", "#,##0", "0.0%"])
    pie(ws, "Content Mix", Reference(ws, min_col=1, min_row=4, max_row=5), Reference(ws, min_col=2, min_row=3, max_row=5), "I3")
    bar(ws, "Titles by release era", Reference(ws, min_col=1, min_row=11, max_row=15),
        Reference(ws, min_col=4, max_col=5, min_row=10, max_row=15), "I21", stacked=True)
    widths(ws, {"A": 18, "B": 10, "C": 10, "D": 11, "E": 11, "F": 10, "G": 11})

    # ---- Genre analysis
    ws = wb.create_sheet("Genre Analysis")
    title(ws, "Top 10 Genres by Content Type", "Genre counts use exact-match COUNTIFS on the comma-separated listed_in field")
    lg, t = R("listed_in"), R("type")
    ex = df.assign(g=df.listed_in.str.split(", ")).explode("g")
    for col0, typ, label, tot in ((1, MOV, "Movies", "'Content Mix'!$B$4"), (5, TV, "TV Shows", "'Content Mix'!$B$5")):
        top = ex[ex.type == typ].g.value_counts().head(10).index.tolist()
        rows = []
        for i, g in enumerate(top):
            r = 5 + i
            a = f"$" + chr(64 + col0) + str(r)
            f = (f'=COUNTIFS({lg},{a},{t},"{typ}")+COUNTIFS({lg},{a}&", *",{t},"{typ}")'
                 f'+COUNTIFS({lg},"*, "&{a},{t},"{typ}")+COUNTIFS({lg},"*, "&{a}&", *",{t},"{typ}")')
            rows.append((g, f, f"={chr(65 + col0)}{r}/{tot}"))
        ws.cell(3, col0, f"Top genres - {label}").font = font(bold=True, size=12)
        table(ws, 4, col0, ["Genre", "Titles", f"% of {label}"], rows, [None, "#,##0", "0.0%"])
        bar(ws, f"Top 10 genres - {label}", Reference(ws, min_col=col0, min_row=5, max_row=14),
            Reference(ws, min_col=col0 + 1, min_row=4, max_row=14), "A17" if col0 == 1 else "A36",
            horizontal=True, w=20, h=9)
    widths(ws, {"A": 28, "B": 10, "C": 12, "D": 3, "E": 28, "F": 10, "G": 14})

    # ---- Insights
    mv_share, tv_share = tot_m / N, tot_t / N
    era = df.assign(e=pd.cut(df.release_year, [1899, 1999, 2009, 2014, 2017, 2021]))
    tv_by_era = era.groupby("e", observed=True).apply(lambda x: (x.type == TV).mean())
    gm = ex[ex.type == MOV].g.value_counts().head(3)
    gt = ex[ex.type == TV].g.value_counts().head(3)
    one_season = ((df.type == TV) & (df.duration_value == 1)).sum() / tot_t
    insights(wb, "Task 1 - Key Observations", [
        ("h", "Portfolio overview"),
        ("t", f"- The catalogue holds {N:,} titles: {tot_m:,} Movies ({pct(mv_share)}) and {tot_t:,} TV Shows ({pct(tv_share)}), i.e. about {tot_m / tot_t:.1f} movies for every TV show."),
        ("t", f"- Movies average {df[df.type == MOV].duration_value.mean():.0f} minutes; TV shows average {df[df.type == TV].duration_value.mean():.1f} seasons and {pct(one_season)} of them have only one season."),
        ("h", "Content mix over time"),
        ("t", f"- TV Shows made up {pct(tv_by_era.iloc[-1])} of titles released in 2018-2021 versus {pct(tv_by_era.iloc[-2])} for 2015-2017, so the mix is moving towards series."),
        ("h", "Genres"),
        ("t", f"- Top movie genres: {', '.join(f'{g} ({n:,})' for g, n in gm.items())}."),
        ("t", f"- Top TV genres: {', '.join(f'{g} ({n:,})' for g, n in gt.items())}."),
        ("t", "- 'International' tags rank near the top for both types, showing Netflix relies heavily on non-US productions to fill the catalogue."),
        ("h", "Business takeaways"),
        ("t", "- Movies still dominate volume, but one-season shows signal many series are short-run or recently launched; a retention-driven strategy would invest in multi-season franchises."),
        ("t", "- Because genres overlap (titles carry several tags), genre counts add up to more than the title count; use them for ranking, not for totals."),
    ])
    data_sheet(wb)
    wb.save(os.path.join(OUT, "Task1_Netflix_Content_Portfolio_Analysis.xlsx"))

    # PNGs
    fig, ax = plt.subplots(1, 2, figsize=(12, 4.5))
    ax[0].pie([tot_m, tot_t], labels=["Movies", "TV Shows"], autopct="%1.1f%%", colors=[TYPE_COL[MOV], TYPE_COL[TV]],
              textprops={"color": "white", "fontweight": "bold"}, labeldistance=1.1)
    ax[0].set_title("Content mix")
    ex2 = ex.groupby(["g", "type"]).size().unstack(fill_value=0)
    top = ex.g.value_counts().head(10).index
    ex2.loc[top[::-1]].plot.barh(stacked=True, ax=ax[1], color=[TYPE_COL[MOV], TYPE_COL[TV]])
    ax[1].set_title("Top 10 genre tags"); ax[1].set_ylabel("")
    png(fig, "task1_content_mix.png")


# ================================================================= TASK 2
def task2():
    wb = new_wb()
    cn, ty, ya, rg = R("country"), R("type"), R("year_added"), R("region")
    vc = known.country.value_counts()
    top15 = vc.head(15).index.tolist()
    ws = wb.create_sheet("Country Ranking")
    title(ws, "Task 2 - Country-wise Content Distribution", "'Not Given' country rows are excluded from share calculations")
    ws["A3"] = "Titles with a known country:"
    ws["A3"].font = font(bold=True)
    ws["D3"] = f'=COUNTA({cn})-COUNTIF({cn},"Not Given")'
    ws["D3"].number_format = "#,##0"
    ws["E3"] = "<- denominator for % share"
    ws["E3"].font = font(italic=True, size=9)
    rows = []
    for i, c in enumerate(top15):
        r = 6 + i
        rows.append((i + 1, c, REGION[c], f"=COUNTIF({cn},B{r})", f'=COUNTIFS({cn},B{r},{ty},"Movie")',
                     f'=COUNTIFS({cn},B{r},{ty},"TV Show")', f"=D{r}/$D$3", f"=SUM($G$6:G{r})", f"=F{r}/D{r}",
                     f'=COUNTIFS({cn},B{r},{ya},">=2019")', f"=J{r}/D{r}",
                     f'=IF(D{r}>=500,"Tier 1 - Core",IF(D{r}>=150,"Tier 2 - Established","Tier 3 - Emerging"))'))
    table(ws, 5, 1, ["Rank", "Country", "Region", "Total Titles", "Movies", "TV Shows", "% of Known Titles",
                     "Cumulative %", "TV Show Mix %", "Added 2019-21", "% Added 2019-21", "Market Tier"], rows,
          [None, None, None, "#,##0", "#,##0", "#,##0", "0.0%", "0.0%", "0.0%", "#,##0", "0.0%", None])
    ws["A22"] = "Top 5 markets combined share"
    ws["A22"].font = font(bold=True)
    ws["G22"] = "=H10"
    ws["G22"].number_format = "0.0%"
    ws["A23"] = "Top 10 markets combined share"
    ws["A23"].font = font(bold=True)
    ws["G23"] = "=H15"
    ws["G23"].number_format = "0.0%"
    bar(ws, "Top 10 countries by titles (Movies vs TV Shows)", Reference(ws, min_col=2, min_row=6, max_row=15),
        Reference(ws, min_col=5, max_col=6, min_row=5, max_row=15), "A26", stacked=True, w=22, h=10)
    widths(ws, {"A": 7, "B": 18, "C": 26, "D": 12, "E": 10, "F": 10, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 22})

    # ---- Regions
    ws = wb.create_sheet("Region Analysis")
    title(ws, "Regional Contribution", "Countries grouped into regions using the mapping in build_analysis.py (see Data!N)")
    regs = known.region.value_counts().index.tolist()
    rows = []
    for i, r_ in enumerate(regs):
        r = 5 + i
        rows.append((r_, f"=COUNTIF({rg},A{r})", f'=COUNTIFS({rg},A{r},{ty},"Movie")', f'=COUNTIFS({rg},A{r},{ty},"TV Show")',
                     f"=B{r}/$B${5 + len(regs)}", f"=D{r}/B{r}", f'=COUNTIFS({rg},A{r},{ya},">=2019")', f"=G{r}/B{r}"))
    last = 5 + len(regs)
    rows.append(("Total (known country)", f"=SUM(B5:B{last - 1})", f"=SUM(C5:C{last - 1})", f"=SUM(D5:D{last - 1})",
                 f"=SUM(E5:E{last - 1})", f"=D{last}/B{last}", f"=SUM(G5:G{last - 1})", f"=G{last}/B{last}"))
    table(ws, 4, 1, ["Region", "Titles", "Movies", "TV Shows", "% of Known Titles", "TV Show Mix %", "Added 2019-21",
                     "% Added 2019-21"], rows, [None, "#,##0", "#,##0", "#,##0", "0.0%", "0.0%", "#,##0", "0.0%"], total_row=True)
    bar(ws, "Titles by region", Reference(ws, min_col=1, min_row=5, max_row=last - 1),
        Reference(ws, min_col=2, min_row=4, max_row=last - 1), "A" + str(last + 3), horizontal=True, w=20, h=9)
    widths(ws, {"A": 28, "B": 10, "C": 10, "D": 10, "E": 14, "F": 14, "G": 14, "H": 14})

    # ---- Expansion
    ws = wb.create_sheet("Expansion Opportunities")
    title(ws, "Emerging Markets: Fastest-Growing Recent Catalogue",
          "Countries with at least 40 titles outside the top-10 markets, ranked by share of titles added 2019-2021")
    rec = known[known.year_added >= 2019].country.value_counts()
    cand = (rec / vc)[(vc >= 40) & (~vc.index.isin(top15[:10]))].sort_values(ascending=False).head(10).index.tolist()
    rows = []
    for i, c in enumerate(cand):
        r = 5 + i
        rows.append((c, REGION[c], f"=COUNTIF({cn},A{r})", f'=COUNTIFS({cn},A{r},{ty},"Movie")',
                     f'=COUNTIFS({cn},A{r},{ty},"TV Show")', f'=COUNTIFS({cn},A{r},{ya},">=2019")', f"=F{r}/C{r}"))
    table(ws, 4, 1, ["Country", "Region", "Total Titles", "Movies", "TV Shows", "Added 2019-21", "% Added 2019-21"], rows,
          [None, None, "#,##0", "#,##0", "#,##0", "#,##0", "0.0%"])
    bar(ws, "% of catalogue added 2019-2021", Reference(ws, min_col=1, min_row=5, max_row=4 + len(cand)),
        Reference(ws, min_col=7, min_row=4, max_row=4 + len(cand)), "A17", w=20, h=9, fmt="0%")
    widths(ws, {"A": 18, "B": 26, "C": 12, "D": 10, "E": 10, "F": 14, "G": 16})

    # ---- Insights
    us, ind = vc["United States"], vc["India"]
    reg_share = (known.region.value_counts() / len(known))
    tvmix = pd.crosstab(known.country, known.type)
    tvmix = (tvmix[TV] / tvmix.sum(axis=1))[top15]
    tv_hi = tvmix.sort_values(ascending=False).head(3)
    af = reg_share.get("Sub-Saharan Africa", 0) + reg_share.get("Middle East & North Africa", 0)
    insights(wb, "Task 2 - Key Observations & Expansion Recommendations", [
        ("h", "Market concentration"),
        ("t", f"- The United States is the largest content market with {us:,} titles ({pct(us / len(known))} of titles with a known country); India is second with {ind:,} ({pct(ind / len(known))})."),
        ("t", f"- The top 5 countries ({', '.join(top15[:5])}) account for {pct(vc.head(5).sum() / len(known))} of titles; the top 10 account for {pct(vc.head(10).sum() / len(known))}."),
        ("t", f"- North America and Europe together hold {pct(reg_share.get('North America', 0) + reg_share.get('Europe', 0))} of titles; South Asia alone holds {pct(reg_share.get('South Asia', 0))}."),
        ("h", "Content-type differences by market"),
        ("t", f"- TV-heavy markets: {', '.join(f'{c} ({pct(v)} TV)' for c, v in tv_hi.items())}. India is movie-led ({pct(1 - tvmix['India'])} movies)."),
        ("h", "Where growth is coming from"),
        ("t", "- " + "; ".join(f"{c} ({pct(rec[c] / vc[c])} of its titles added 2019-21)" for c in cand[:5]) + "."),
        ("t", f"- Sub-Saharan Africa plus Middle East & North Africa contribute only {pct(af)} of titles, but several of their markets (Egypt, Nigeria, South Africa) show very recent, fast catalogue build-up."),
        ("h", "Recommendations"),
        ("t", "1. Defend the core: keep investing in US, UK and India where the catalogue is deepest and customer base is established."),
        ("t", "2. Scale emerging markets with local-language originals: Nigeria, South Africa, Egypt, Philippines and Indonesia grew mostly in the last three years and have room to deepen."),
        ("t", f"3. Treat TV-led markets (Pakistan, South Korea, Japan) as series hubs: build multi-season franchises and export them internationally."),
        ("t", "4. Under-represented regions (Latin America outside Mexico/Brazil, Sub-Saharan Africa, Oceania beyond Australia) are low-cost test beds for licensing before heavy originals spend."),
        ("t", "Caveat: this dataset measures catalogue availability by production country, not viewing hours or subscriber revenue, so treat 'performance' as catalogue contribution."),
    ])
    data_sheet(wb)
    wb.save(os.path.join(OUT, "Task2_Global_Market_Performance_Analysis.xlsx"))

    fig, ax = plt.subplots(1, 2, figsize=(14, 5))
    ct = pd.crosstab(known.country, known.type).loc[top15[:10]]
    ct.plot.bar(stacked=True, ax=ax[0], color=[TYPE_COL[MOV], TYPE_COL[TV]])
    ax[0].set_title("Top 10 countries by titles"); ax[0].set_xlabel("")
    known.region.value_counts()[::-1].plot.barh(ax=ax[1], color="#B20710")
    ax[1].set_title("Titles by region")
    png(fig, "task2_country_region.png")


# ================================================================= TASK 3
def task3():
    wb = new_wb()
    rt, ty, sg, ya = R("rating"), R("type"), R("audience_segment"), R("year_added")
    ws = wb.create_sheet("Rating Analysis")
    title(ws, "Task 3 - Content Ratings by Type", "Each rating is mapped to an audience segment (column B)")
    ratings = df.rating.value_counts().index.tolist()
    n = len(ratings)
    rows = []
    for i, r_ in enumerate(ratings):
        r = 5 + i
        rows.append((r_, SEGMENT[r_], f"=COUNTIF({rt},A{r})", f'=COUNTIFS({rt},A{r},{ty},"Movie")',
                     f'=COUNTIFS({rt},A{r},{ty},"TV Show")', f"=D{r}/D${5 + n}", f"=E{r}/E${5 + n}", f"=C{r}/C${5 + n}"))
    L_ = 5 + n
    rows.append(("Total", "", f"=SUM(C5:C{L_ - 1})", f"=SUM(D5:D{L_ - 1})", f"=SUM(E5:E{L_ - 1})",
                 f"=SUM(F5:F{L_ - 1})", f"=SUM(G5:G{L_ - 1})", f"=SUM(H5:H{L_ - 1})"))
    table(ws, 4, 1, ["Rating", "Audience Segment", "Total", "Movies", "TV Shows", "% of Movies", "% of TV Shows", "% of All"],
          rows, [None, None, "#,##0", "#,##0", "#,##0", "0.0%", "0.0%", "0.0%"], total_row=True)
    bar(ws, "Rating mix: Movies vs TV Shows (% of each type)", Reference(ws, min_col=1, min_row=5, max_row=L_ - 1),
        Reference(ws, min_col=6, max_col=7, min_row=4, max_row=L_ - 1), f"A{L_ + 3}", w=24, h=10, fmt="0%")
    widths(ws, {"A": 12, "B": 26, "C": 10, "D": 10, "E": 10, "F": 12, "G": 14, "H": 10})

    # segments
    ws = wb.create_sheet("Audience Segments")
    title(ws, "Audience Segments", "Kids = TV-Y/Y7/Y7-FV/G/TV-G | Family = TV-PG/PG | Teens = TV-14/PG-13 | Adults = TV-MA/R/NC-17 | Unrated = NR/UR")
    rows = []
    for i, s in enumerate(SEG_ORDER):
        r = 5 + i
        rows.append((s, f"=COUNTIF({sg},A{r})", f'=COUNTIFS({sg},A{r},{ty},"Movie")', f'=COUNTIFS({sg},A{r},{ty},"TV Show")',
                     f"=B{r}/B$10", f"=C{r}/C$10", f"=D{r}/D$10", f"=G{r}-F{r}"))
    rows.append(("Total", "=SUM(B5:B9)", "=SUM(C5:C9)", "=SUM(D5:D9)", "=SUM(E5:E9)", "=SUM(F5:F9)", "=SUM(G5:G9)", ""))
    table(ws, 4, 1, ["Segment", "Total", "Movies", "TV Shows", "% of All", "% of Movies", "% of TV Shows", "TV minus Movie (pp)"],
          rows, [None, "#,##0", "#,##0", "#,##0", "0.0%", "0.0%", "0.0%", "0.0%"], total_row=True)
    pie(ws, "Audience segments (all titles)", Reference(ws, min_col=1, min_row=5, max_row=9),
        Reference(ws, min_col=2, min_row=4, max_row=9), "J3")
    bar(ws, "Segments: Movies vs TV Shows", Reference(ws, min_col=1, min_row=5, max_row=9),
        Reference(ws, min_col=6, max_col=7, min_row=4, max_row=9), "J22", w=16, h=9, fmt="0%")
    ws["A13"] = "Segment share of titles added each year"
    ws["A13"].font = font(bold=True, size=12)
    years = list(range(2016, 2022))
    rows = []
    for i, y in enumerate(years):
        r = 15 + i
        rows.append([y] + [f'=COUNTIFS({sg},"{s}",{ya},$A{r})/COUNTIF({ya},$A{r})' for s in SEG_ORDER])
    table(ws, 14, 1, ["Year added"] + SEG_ORDER, rows, ["0"] + ["0.0%"] * 5)
    line(ws, "Segment share of yearly additions", Reference(ws, min_col=1, min_row=15, max_row=20),
         Reference(ws, min_col=2, max_col=6, min_row=14, max_row=20), "A23", w=20, h=9, fmt="0%")
    widths(ws, {"A": 26, "B": 12, "C": 14, "D": 12, "E": 12, "F": 14, "G": 16, "H": 20})

    # insights
    seg = df.audience_segment.value_counts(normalize=True)
    sm = df[df.type == MOV].audience_segment.value_counts(normalize=True)
    st = df[df.type == TV].audience_segment.value_counts(normalize=True)
    top_r = df.rating.value_counts(normalize=True)
    yr = pd.crosstab(df.year_added, df.audience_segment, normalize="index")
    insights(wb, "Task 3 - Audience Insights & Recommendations", [
        ("h", "Dominant segments"),
        ("t", f"- Adults (TV-MA, R, NC-17) is the largest segment at {pct(seg['Adults'])} of titles, followed by Teens at {pct(seg['Teens'])}. Together they cover {pct(seg['Adults'] + seg['Teens'])} of the catalogue."),
        ("t", f"- TV-MA ({pct(top_r['TV-MA'])}) and TV-14 ({pct(top_r['TV-14'])}) are the two most common individual ratings."),
        ("t", f"- Kids content is {pct(seg['Kids'])} and Family (Parental Guidance) {pct(seg['Family (Parental Guidance)'])} of titles - a much smaller family-viewing offer."),
        ("h", "Movies vs TV Shows"),
        ("t", f"- Movies: Adults {pct(sm['Adults'])}, Teens {pct(sm['Teens'])}, Kids {pct(sm['Kids'])}. TV Shows: Adults {pct(st['Adults'])}, Teens {pct(st['Teens'])}, Kids {pct(st['Kids'])}."),
        ("t", f"- TV Shows skew towards Kids content ({pct(st['Kids'])} vs {pct(sm['Kids'])} of movies) and carry a higher TV-MA share ({pct(((df.type == TV) & (df.rating == 'TV-MA')).sum() / tot_t)} vs {pct(((df.type == MOV) & (df.rating == 'TV-MA')).sum() / tot_m)}). Movies hold almost all R, PG-13 and PG titles, which TV rarely uses, so movies lead on the Adults and Teens segments."),
        ("h", "Trend"),
        ("t", f"- Adult share of yearly additions moved from {pct(yr.loc[2016, 'Adults'])} (2016) to {pct(yr.loc[2021, 'Adults'])} (2021); Kids share moved from {pct(yr.loc[2016, 'Kids'])} to {pct(yr.loc[2021, 'Kids'])}."),
        ("h", "Audience-focused recommendations"),
        ("t", "1. Keep adult drama/thriller output as the retention engine but diversify with teen-friendly (TV-14/PG-13) titles, the second-largest and broadest-appeal segment."),
        ("t", "2. Grow family and kids originals, especially series: this segment is under-served relative to its role in household sign-ups."),
        ("t", "3. Use ratings as a localisation filter: promote TV-14/PG-13 content in conservative markets and TV-MA in mature-content-friendly markets."),
        ("t", "4. Monitor the 'Unrated' bucket (NR/UR) - small but a data-quality and compliance risk; assign formal ratings."),
    ])
    data_sheet(wb)
    wb.save(os.path.join(OUT, "Task3_Audience_Segmentation_Analysis.xlsx"))

    fig, ax = plt.subplots(1, 2, figsize=(14, 5))
    ct = pd.crosstab(df.rating, df.type, normalize="columns").loc[ratings]
    ct.plot.bar(ax=ax[0], color=[TYPE_COL[MOV], TYPE_COL[TV]])
    ax[0].set_title("Rating mix (% of each type)"); ax[0].yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1)); ax[0].set_xlabel("")
    seg_t = pd.crosstab(df.audience_segment, df.type, normalize="columns").loc[SEG_ORDER]
    seg_t.plot.bar(ax=ax[1], color=[TYPE_COL[MOV], TYPE_COL[TV]])
    ax[1].set_title("Audience segment share by type"); ax[1].yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1)); ax[1].set_xlabel("")
    plt.setp(ax[1].get_xticklabels(), rotation=20)
    png(fig, "task3_audience_segments.png")


# ================================================================= TASK 4
def task4():
    wb = new_wb()
    ty, ya, ry = R("type"), R("year_added"), R("release_year")
    # additions trend
    ws = wb.create_sheet("Additions Trend")
    title(ws, "Titles Added to Netflix per Year", "2021 is partial: data ends 25-Sep-2021")
    years = list(range(2008, 2022))
    rows = []
    for i, y in enumerate(years):
        r = 5 + i
        rows.append((y, f'=COUNTIFS({ya},A{r},{ty},"Movie")', f'=COUNTIFS({ya},A{r},{ty},"TV Show")', f"=B{r}+C{r}",
                     "" if i == 0 else f'=IFERROR((D{r}-D{r - 1})/D{r - 1},"")', f"=SUM($D$5:D{r})", f"=C{r}/D{r}"))
    table(ws, 4, 1, ["Year Added", "Movies", "TV Shows", "Total", "YoY Growth", "Cumulative Titles", "TV Show Share"], rows,
          [None, "#,##0", "#,##0", "#,##0", "0.0%", "#,##0", "0.0%"])
    last = 4 + len(years)
    line(ws, "Titles added per year", Reference(ws, min_col=1, min_row=5, max_row=last),
         Reference(ws, min_col=2, max_col=4, min_row=4, max_row=last), "I3", w=20, h=9)
    bar(ws, "Cumulative catalogue size", Reference(ws, min_col=1, min_row=5, max_row=last),
        Reference(ws, min_col=6, min_row=4, max_row=last), "I22", w=20, h=9)
    widths(ws, {"A": 12, "B": 10, "C": 10, "D": 10, "E": 12, "F": 16, "G": 14})
    # rows: 2008->5 ... 2016->13, 2019->16, 2021->18
    # release trend
    ws = wb.create_sheet("Release Year Trend")
    title(ws, "Titles by Original Release Year", "Pre-2000 titles are grouped in the last row")
    rows = []
    for i, y in enumerate(range(2000, 2022)):
        r = 5 + i
        rows.append((y, f'=COUNTIFS({ry},A{r},{ty},"Movie")', f'=COUNTIFS({ry},A{r},{ty},"TV Show")', f"=B{r}+C{r}",
                     f'=IFERROR((D{r}-D{r - 1})/D{r - 1},"")' if i else ""))
    rows.append(("Before 2000", f'=COUNTIFS({ry},"<2000",{ty},"Movie")', f'=COUNTIFS({ry},"<2000",{ty},"TV Show")', "=B27+C27", ""))
    table(ws, 4, 1, ["Release Year", "Movies", "TV Shows", "Total", "YoY Change"], rows, [None, "#,##0", "#,##0", "#,##0", "0.0%"])
    line(ws, "Titles by release year (2000-2021)", Reference(ws, min_col=1, min_row=5, max_row=26),
         Reference(ws, min_col=2, max_col=3, min_row=4, max_row=26), "G3", w=20, h=9)
    widths(ws, {"A": 14, "B": 10, "C": 10, "D": 10, "E": 12})
    # growth phases
    ws = wb.create_sheet("Growth Phases")
    title(ws, "Growth Periods", "Phases defined from the additions curve; counts pull from 'Additions Trend'")
    ph = [("Pilot / early catalogue", 2008, 2014), ("Rapid expansion", 2015, 2018), ("Scale & peak output", 2019, 2020), ("2021 year-to-date (partial)", 2021, 2021)]
    rows = []
    A = "'Additions Trend'!"
    for i, (p, a, b) in enumerate(ph):
        r = 5 + i
        rows.append((p, a, b, f"=C{r}-B{r}+1", f'=SUMIFS({A}$D$5:$D$18,{A}$A$5:$A$18,">="&B{r},{A}$A$5:$A$18,"<="&C{r})',
                     f"=E{r}/D{r}", f'=SUMIFS({A}$C$5:$C$18,{A}$A$5:$A$18,">="&B{r},{A}$A$5:$A$18,"<="&C{r})/E{r}'))
    table(ws, 4, 1, ["Phase", "From", "To", "Years", "Titles Added", "Avg per Year", "TV Show Share"], rows,
          [None, "0", "0", "0", "#,##0", "#,##0", "0.0%"])
    ws["A11"] = "Growth metrics"
    ws["A11"].font = font(bold=True, size=12)
    table(ws, 12, 1, ["Metric", "Value", "Note"], [
        ("CAGR of yearly additions 2016-2019", f"=({A}D16/{A}D13)^(1/3)-1", "(2019 additions / 2016 additions)^(1/3) - 1"),
        ("Peak addition year", f"=INDEX({A}A5:A17,MATCH(MAX({A}D5:D17),{A}D5:D17,0))", "Full years only (2008-2020)"),
        ("Titles added in peak year", f"=MAX({A}D5:D17)", ""),
        ("TV Show share of additions 2016", f"={A}G13", ""),
        ("TV Show share of additions 2020", f"={A}G17", ""),
        ("Change in TV share 2016 to 2020 (pp)", "=B17-B16", "Positive = shift towards series"),
    ], [None, None, None])
    ws["B13"].number_format = "0.0%"; ws["B15"].number_format = "#,##0"; ws["B16"].number_format = "0.0%"
    ws["B17"].number_format = "0.0%"; ws["B18"].number_format = "0.0%"; ws["B14"].number_format = "0"
    widths(ws, {"A": 38, "B": 14, "C": 44, "D": 8, "E": 14, "F": 14, "G": 14})

    # dashboard
    ws = wb.create_sheet("Dashboard", 0)
    title(ws, "Task 4 - Content Growth & Trend Dashboard", "Executive view of Netflix catalogue growth (live formulas)")
    tiles = [("Total titles", "=COUNTA(" + R("show_id") + ")", "#,##0"),
             ("Peak yearly additions", "='Growth Phases'!B15", "#,##0"),
             ("Peak year", "='Growth Phases'!B14", "0"),
             ("CAGR 2016-19", "='Growth Phases'!B13", "0.0%"),
             ("TV share 2016", "='Growth Phases'!B16", "0.0%"),
             ("TV share 2020", "='Growth Phases'!B17", "0.0%")]
    for i, (lab, f_, fm) in enumerate(tiles):
        col = 1 + i * 2
        ws.cell(4, col, lab).font = font(bold=True, color="FFFFFF", size=9)
        ws.cell(4, col).fill = PatternFill("solid", fgColor=RED)
        ws.cell(4, col).alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        c = ws.cell(5, col, f_)
        c.font = font(bold=True, size=18, color=DARK)
        c.number_format = fm
        c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill("solid", fgColor=GREY)
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
    ws.row_dimensions[5].height = 34
    for j in range(1, 13):
        ws.column_dimensions[chr(64 + j)].width = 11
    a = wb["Additions Trend"]
    line(ws, "Titles added per year", Reference(a, min_col=1, min_row=5, max_row=18),
         Reference(a, min_col=2, max_col=4, min_row=4, max_row=18), "A8", w=17, h=9)
    bar(ws, "TV Show share of yearly additions", Reference(a, min_col=1, min_row=5, max_row=18),
        Reference(a, min_col=7, min_row=4, max_row=18), "H8", w=17, h=9, fmt="0%")
    r_ = wb["Release Year Trend"]
    line(ws, "Titles by release year", Reference(r_, min_col=1, min_row=5, max_row=26),
         Reference(r_, min_col=2, max_col=3, min_row=4, max_row=26), "A27", w=17, h=9)
    bar(ws, "Cumulative catalogue size", Reference(a, min_col=1, min_row=5, max_row=18),
        Reference(a, min_col=6, min_row=4, max_row=18), "H27", w=17, h=9)

    # insights
    add = df.groupby(["year_added", "type"]).size().unstack(fill_value=0)
    add["T"] = add.sum(axis=1)
    peak = int(add.loc[:2020, "T"].idxmax())
    cagr = (add.loc[2019, "T"] / add.loc[2016, "T"]) ** (1 / 3) - 1
    rel = df.groupby(["release_year", "type"]).size().unstack(fill_value=0)
    mpeak, tpeak = int(rel[MOV].loc[:2020].idxmax()), int(rel[TV].loc[:2020].idxmax())
    insights(wb, "Task 4 - Strategic Insights", [
        ("h", "Growth periods"),
        ("t", f"- 2008-2014: almost no activity ({int(add.loc[:2014, 'T'].sum())} titles in seven years) - Netflix was still a pilot catalogue for this dataset."),
        ("t", f"- 2015-2018: rapid expansion. Yearly additions rose from {int(add.loc[2015, 'T'])} (2015) to {int(add.loc[2018, 'T']):,} (2018); 2016 to 2019 CAGR was {pct(cagr)}."),
        ("t", f"- 2019-2020: scale/peak. Additions peaked at {int(add.loc[peak, 'T']):,} in {peak} and stayed near that level in 2020 ({int(add.loc[2020, 'T']):,}). 2021 is partial ({int(add.loc[2021, 'T']):,} by 25 Sep) so it should not be read as a decline."),
        ("h", "Content-type shift"),
        ("t", f"- TV Shows were {pct(add.loc[2016, TV] / add.loc[2016, 'T'])} of additions in 2016 but {pct(add.loc[2020, TV] / add.loc[2020, 'T'])} in 2020 ({pct(add.loc[2021, TV] / add.loc[2021, 'T'])} in 2021 YTD): the additions pipeline stayed movie-heavy as volume scaled."),
        ("t", f"- By original release year the picture differs: TV was {pct(rel.loc[2021, TV] / (rel.loc[2021, TV] + rel.loc[2021, MOV]))} of 2021 releases versus {pct(rel.loc[2016, TV] / (rel.loc[2016, TV] + rel.loc[2016, MOV]))} of 2016 releases. Movie releases peaked in {mpeak} ({int(rel[MOV].loc[mpeak])}), TV releases in {tpeak} ({int(rel[TV].loc[tpeak])}), so newer content is tilting towards series."),
        ("h", "Strategic recommendations"),
        ("t", "1. Plan content investment around the 2015-2019 growth curve: volume growth has plateaued, so future growth should come from quality, franchise depth and new markets rather than raw title count."),
        ("t", "2. Rebalance towards series: new releases are TV-heavier than the additions pipeline, so closing that gap (more multi-season originals) supports repeat viewing and retention."),
        ("t", "3. Forecast with caution: with 2021 partial and 2020 potentially affected by production disruption, use 2016-2019 as the clean growth baseline."),
        ("t", f"4. Note the backlog effect: {pct((df[df.year_added >= 2019].year_added - df[df.year_added >= 2019].release_year >= 2).mean())} of titles added since 2019 were released two or more years earlier, i.e. growth came from licensing back-catalogue as well as new production."),
    ])
    data_sheet(wb)
    wb.save(os.path.join(OUT, "Task4_Content_Growth_Trend_Analysis.xlsx"))

    fig, ax = plt.subplots(1, 2, figsize=(14, 5))
    add[[MOV, TV]].plot.bar(stacked=True, ax=ax[0], color=[TYPE_COL[MOV], TYPE_COL[TV]])
    ax[0].set_title("Titles added per year (2021 partial)"); ax[0].set_xlabel("")
    rel2 = rel.loc[2000:]
    rel2[[MOV, TV]].plot(ax=ax[1], color=[TYPE_COL[MOV], TYPE_COL[TV]], marker="o")
    ax[1].set_title("Titles by release year"); ax[1].set_xlabel("")
    png(fig, "task4_growth_trend.png")


if __name__ == "__main__":
    task1(); task2(); task3(); task4()
    print("done", N)
