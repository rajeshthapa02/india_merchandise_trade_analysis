#!/usr/bin/env python
# coding: utf-8

# In[41]:


from pathlib import Path
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
try:
    from scipy import stats
except ImportError:
    stats = None
from matplotlib.ticker import FuncFormatter


# In[43]:


WORKBOOK = Path(r"D:\India_Merchandise_Trade_2011-2026.xlsx")

OUT = Path(r"D:\trade_project_outputs")
START_FY, END_FY = 2011, 2025


# In[45]:


EXCLUDE_OTHER_COMMODITIES = False


# In[47]:


def clean_label(text):
    """Remove source numbering such as '2.', 'IV.' or 'A.' from category names."""
    t = re.sub(r"^\s*(\d+|[IVX]+|[A-D])\.\s*", "", str(text)).strip()
    return re.sub(r"\s+", " ", t.replace("U S A", "USA").replace("U A E", "UAE")
                  .replace("U K", "UK").replace("O E C D", "OECD").replace("O P E C", "OPEC")
                  .replace("S A A R C", "SAARC"))

def thousands(axis):
    axis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}"))

def read_sheet(sheet, header=5):
    return pd.read_excel(WORKBOOK, sheet_name=sheet, header=header)

def parse_fy(value):
    if pd.isna(value):
        return np.nan
    match = re.search(r"(20\d{2})", str(value).strip())
    return int(match.group(1)) if match else np.nan

def numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip()
        .replace({"-": np.nan, "—": np.nan, "nan": np.nan, "": np.nan}),
        errors="coerce"
    )

def format_years(df):
    df = df.copy()
    df["fy_start"] = df["Year"].apply(parse_fy)
    return df[df["fy_start"].between(START_FY, END_FY)].copy()

def save_chart(filename, title, xlabel, ylabel, note=None):
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.xticks(rotation=45, ha="right")
    plt.grid(axis="y", alpha=0.25)
    if note:
        plt.figtext(0.01, 0.01, note, ha="left", fontsize=8)
    plt.tight_layout(rect=(0, 0.04, 1, 1))
    plt.savefig(OUT / filename, dpi=250, bbox_inches="tight")
    plt.close()


def line_chart(x, series, title, ylabel, filename):
    plt.figure(figsize=(11, 6))
    for y, label in series:
        plt.plot(x, y, marker="o", linewidth=1.8, label=label)
    plt.title(title)
    plt.xlabel("Financial year")
    plt.ylabel(ylabel)
    if "₹" in ylabel or "US$" in ylabel:
        thousands(plt.gca().yaxis)
    plt.xticks(rotation=45, ha="right")
    plt.grid(True, alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(OUT / filename, dpi=250)
    plt.close()


# In[49]:


def bar_chart(labels, values, title, ylabel, filename, horizontal=False):
    plt.figure(figsize=(11, 6.5))
    if horizontal:
        order = np.argsort(values)
        plt.barh(np.array(labels)[order], np.array(values)[order])
        plt.xlabel(ylabel)
        plt.ylabel("Category")
        if "₹" in ylabel:
            thousands(plt.gca().xaxis)
        plt.title(title)
        plt.grid(axis="x", alpha=0.25)
        plt.tight_layout()
        plt.savefig(OUT / filename, dpi=250, bbox_inches="tight")
        plt.close()
    else:
        plt.bar(labels, values)
        if "₹" in ylabel or "US$" in ylabel:
            thousands(plt.gca().yaxis)
        save_chart(filename, title, "Financial year" if str(list(labels)[0])[:2] == "20" else "Category", ylabel)

def monthly_annual(sheet, unit):
    df = read_sheet(sheet)
    
    df = df.rename(columns={df.columns[1]: "Year", df.columns[2]: "Month",
                            df.columns[3]: "Exports", df.columns[4]: "Imports"})
    df = format_years(df)
    df["Exports"] = numeric(df["Exports"])
    df["Imports"] = numeric(df["Imports"])
    annual = df.groupby("fy_start", as_index=False).agg(
        exports=("Exports", "sum"), imports=("Imports", "sum"),
        months=("Month", "count")
    )
    annual = annual[annual["months"] >= 12].copy()
    annual["financial_year"] = annual["fy_start"].astype(int).map(
        lambda y: f"{y}-{(y + 1) % 100:02d}"
    )
    annual["trade_balance"] = annual["exports"] - annual["imports"]
    annual["export_yoy_pct"] = annual["exports"].pct_change() * 100
    annual["import_yoy_pct"] = annual["imports"].pct_change() * 100
    annual["coverage_pct"] = annual["exports"] / annual["imports"] * 100
    annual["unit"] = unit
    return annual


# In[51]:


def oil_annual(sheet):
    
    df = pd.read_excel(WORKBOOK, sheet_name=sheet, header=5)
    df = df.rename(columns={df.columns[1]: "Year", df.columns[2]: "Month"})
    
    df = df.rename(columns={df.columns[3]: "export_oil", df.columns[4]: "export_non_oil",
                            df.columns[5]: "import_oil", df.columns[6]: "import_non_oil"})
    df = format_years(df)
    for col in ["export_oil", "export_non_oil", "import_oil", "import_non_oil"]:
        df[col] = numeric(df[col])
    annual = df.groupby("fy_start", as_index=False).agg(
        export_oil=("export_oil", "sum"), export_non_oil=("export_non_oil", "sum"),
        import_oil=("import_oil", "sum"), import_non_oil=("import_non_oil", "sum"),
        months=("Month", "count")
    )
    annual = annual[annual["months"] >= 12].copy()
    annual["financial_year"] = annual["fy_start"].astype(int).map(
        lambda y: f"{y}-{(y + 1) % 100:02d}"
    )
    annual["export_total"] = annual["export_oil"] + annual["export_non_oil"]
    annual["import_total"] = annual["import_oil"] + annual["import_non_oil"]
    annual["oil_export_share"] = annual["export_oil"] / annual["export_total"] * 100
    annual["oil_import_share"] = annual["import_oil"] / annual["import_total"] * 100
    annual["non_oil_export_share"] = 100 - annual["oil_export_share"]
    annual["non_oil_import_share"] = 100 - annual["oil_import_share"]
    return annual


# In[53]:


def commodity_table(sheet):
    df = read_sheet(sheet)
    year_col = df.columns[1]
    df = df.rename(columns={year_col: "Year"})
    
    df = df[df["Year"].astype(str).str.strip().str.match(r"^20\d{2}-\d{2}$")].copy()
   
    category_cols = [c for c in df.columns[2:] if not str(c).lower().startswith("unnamed")]
    category_cols = [c for c in category_cols if not str(c).lower().startswith("total")]
    df = format_years(df)
    for c in category_cols:
        df[c] = numeric(df[c])
    df["financial_year"] = df["fy_start"].astype(int).map(
        lambda y: f"{y}-{(y + 1) % 100:02d}"
    )
    return df, category_cols

def direction_table(sheet):
    df = read_sheet(sheet)
    df = df.rename(columns={df.columns[1]: "Year", df.columns[2]: "Flow"})
    df["Year"] = df["Year"].ffill()
    df = format_years(df)
    category_cols = [c for c in df.columns[3:] if not str(c).lower().startswith("unnamed")]
    for c in category_cols:
        df[c] = numeric(df[c])
    df["Flow"] = df["Flow"].astype(str).str.strip().str.lower()
    return df, category_cols


def run_hypothesis_tests(rupees):
    """H1: equal mean YoY growth of imports and exports; H2: no linear trend in the deficit."""
    g = rupees.dropna(subset=["export_yoy_pct"])
    t_res = stats.ttest_rel(g["import_yoy_pct"], g["export_yoy_pct"])
    w_res = stats.wilcoxon(g["import_yoy_pct"], g["export_yoy_pct"])
    years = np.arange(len(rupees))
    lr = stats.linregress(years, rupees["trade_balance"])
    n_years = len(rupees) - 1
    cagr_x = ((rupees.exports.iloc[-1] / rupees.exports.iloc[0]) ** (1 / n_years) - 1) * 100
    cagr_m = ((rupees.imports.iloc[-1] / rupees.imports.iloc[0]) ** (1 / n_years) - 1) * 100
    out = pd.DataFrame([
        ["H1 paired t-test (import vs export YoY growth)", f"mean imports={g.import_yoy_pct.mean():.2f}%, mean exports={g.export_yoy_pct.mean():.2f}%",
         f"t({len(g)-1})={t_res.statistic:.2f}", round(t_res.pvalue, 3), "Fail to reject H0" if t_res.pvalue >= 0.05 else "Reject H0"],
        ["H1 Wilcoxon signed-rank (robustness)", "", f"W={w_res.statistic:.1f}", round(w_res.pvalue, 3), "Fail to reject H0" if w_res.pvalue >= 0.05 else "Reject H0"],
        ["H2 linear trend in trade balance (Rs crore per year)", f"slope={lr.slope:,.0f}, R2={lr.rvalue**2:.2f}",
         f"t({len(rupees)-2})={lr.slope/lr.stderr:.2f}", float(f"{lr.pvalue:.5f}"), "Reject H0" if lr.pvalue < 0.05 else "Fail to reject H0"],
        ["Compound annual growth (FY first to last)", f"exports={cagr_x:.2f}%, imports={cagr_m:.2f}%", "", "", "descriptive"],
    ], columns=["test", "estimate", "statistic", "p_value", "decision_at_5pct"])
    out.to_csv(OUT / "hypothesis_tests.csv", index=False)
    print(out.to_string(index=False))


# In[55]:


def main():
    if not WORKBOOK.exists():
        raise FileNotFoundError(f"Workbook not found: {WORKBOOK}. Check that the file is saved exactly at that location.")
    OUT.mkdir(exist_ok=True)
    for stale in OUT.iterdir():
        if stale.is_file() and stale.suffix.lower() in {".png", ".csv"}:
            stale.unlink()


    rupees = monthly_annual("Monthly_Trade_Rupees", "₹ crore")
    usd = monthly_annual("Monthly_Trade_USD", "US$ million")
    rupees.to_csv(OUT / "annual_summary_rupees.csv", index=False)
    usd.to_csv(OUT / "annual_summary_usd.csv", index=False)


    line_chart(rupees.financial_year, [(rupees.exports, "Exports"), (rupees.imports, "Imports")],
               "Annual merchandise exports and imports (rupees)", "₹ crore", "01_annual_trade_rupees.png")
    line_chart(usd.financial_year, [(usd.exports, "Exports"), (usd.imports, "Imports")],
               "Annual merchandise exports and imports (US dollars)", "US$ million", "02_annual_trade_usd.png")
    bar_chart(rupees.financial_year, rupees.trade_balance, "Merchandise trade balance (rupees)",
              "Exports minus imports (₹ crore)", "03_trade_balance_rupees.png")
    bar_chart(usd.financial_year, usd.trade_balance, "Merchandise trade balance (US dollars)",
              "Exports minus imports (US$ million)", "04_trade_balance_usd.png")
    recent = rupees.tail(8)
    bar_chart(recent.financial_year, recent.trade_balance, "Trade balance: most recent eight complete years",
              "Trade balance (₹ crore)", "05_recent_trade_balance.png")
    line_chart(rupees.financial_year, [(rupees.export_yoy_pct, "Exports"), (rupees.import_yoy_pct, "Imports")],
               "Year-on-year growth in trade values (rupees)", "Growth (%)", "06_growth_rupees.png")
    line_chart(usd.financial_year, [(usd.export_yoy_pct, "Exports"), (usd.import_yoy_pct, "Imports")],
               "Year-on-year growth in trade values (US dollars)", "Growth (%)", "07_growth_usd.png")
    line_chart(rupees.financial_year, [(rupees.coverage_pct, "Export-import coverage")],
               "Export-import coverage ratio", "Exports / imports (%)", "08_coverage_ratio.png")

    rupee_base = rupees.exports.iloc[0]; usd_base = usd.exports.iloc[0]
    idx_rupee = rupees.exports / rupee_base * 100 if rupee_base else rupees.exports * np.nan
    idx_usd = usd.exports / usd_base * 100 if usd_base else usd.exports * np.nan
    line_chart(rupees.financial_year, [(idx_rupee, "Exports, rupees"), (idx_usd, "Exports, US dollars")],
               "Export-value indices (FY 2011-12 = 100)", "Index (FY 2011-12 = 100)", "09_export_value_indices.png")
    try:
        run_hypothesis_tests(rupees)
    except Exception as exc:
        print(f"WARNING: hypothesis tests skipped ({type(exc).__name__}: {exc})")


    oilr = oil_annual("Oil_NonOil_Rupees")
    oilu = oil_annual("Oil_NonOil_USD")
    oilr.to_csv(OUT / "annual_oil_non_oil_rupees.csv", index=False)
    oilu.to_csv(OUT / "annual_oil_non_oil_usd.csv", index=False)
    line_chart(oilr.financial_year, [(oilr.export_oil, "Oil exports"), (oilr.export_non_oil, "Non-oil exports")],
               "Oil and non-oil exports (rupees)", "₹ crore", "10_oil_non_oil_exports.png")
    line_chart(oilr.financial_year, [(oilr.import_oil, "Oil imports"), (oilr.import_non_oil, "Non-oil imports")],
               "Oil and non-oil imports (rupees)", "₹ crore", "11_oil_non_oil_imports.png")
    line_chart(oilr.financial_year, [(oilr.oil_export_share, "Oil export share"), (oilr.oil_import_share, "Oil import share")],
               "Oil share of total exports and total imports", "Share (%)", "12_oil_shares.png")
    line_chart(oilr.financial_year, [(oilr.non_oil_export_share, "Non-oil export share"),
                                    (oilr.non_oil_import_share, "Non-oil import share")],
               "Non-oil share of total exports and total imports", "Share (%)", "13_non_oil_shares.png")
    latest = oilr.iloc[-1]
    bar_chart(["Oil exports", "Non-oil exports", "Oil imports", "Non-oil imports"],
              [latest.export_oil, latest.export_non_oil, latest.import_oil, latest.import_non_oil],
              f"Oil and non-oil trade components (FY {latest.financial_year})", "₹ crore", "14_latest_oil_components.png")
    oilr_base = oilr.export_oil.iloc[0]; oilu_base = oilu.export_oil.iloc[0]
    oil_idx_r = oilr.export_oil / oilr_base * 100 if oilr_base else oilr.export_oil * np.nan
    oil_idx_u = oilu.export_oil / oilu_base * 100 if oilu_base else oilu.export_oil * np.nan
    line_chart(oilr.financial_year, [(oil_idx_r, "Oil exports, rupees"), (oil_idx_u, "Oil exports, US dollars")],
               "Oil-export indices (FY 2011-12 = 100)", "Index (FY 2011-12 = 100)", "15_oil_export_indices.png")
    line_chart(oilr.financial_year,
               [(oilr.export_oil-oilr.import_oil, "Oil balance"),
                (oilr.export_non_oil-oilr.import_non_oil, "Non-oil balance")],
               "Oil and non-oil merchandise trade balances", "₹ crore", "16_oil_non_oil_balances.png")
    exp_df, exp_cols = commodity_table("Exports_Commodities")
    imp_df, imp_cols = commodity_table("Imports_Commodities_2")
    for offset, (df, cols, label, prefix) in enumerate([
        (exp_df, exp_cols, "Export", "export"), (imp_df, imp_cols, "Import", "import")
    ]):
        last = df.sort_values("fy_start").iloc[-1]
        plot_cols = [c for c in cols if not (EXCLUDE_OTHER_COMMODITIES and "other commodities" in str(c).lower())]
        vals = pd.Series({str(c): last[c] for c in plot_cols}).dropna().sort_values(ascending=False).head(8)
        level_fig = 17 if prefix == "export" else 21
        share_fig = 18 if prefix == "export" else 22
        change_fig = 19 if prefix == "export" else 23
        trend_fig = 20 if prefix == "export" else 24
        bar_chart(vals.index, vals.values, f"Eight largest {label.lower()} commodity groups (FY {last.financial_year})",
                  "₹ crore", f"{level_fig:02d}_top_{prefix}_commodities.png", horizontal=True)
        total_col = next((c for c in df.columns if str(c).strip().lower() in ("total exports", "total imports")), None)
        totals = numeric(df[total_col]) if total_col is not None else df[cols].sum(axis=1)
        shares = df[cols].div(totals.replace(0, np.nan), axis=0) * 100
        latest_shares = shares.loc[last.name, plot_cols].sort_values(ascending=False).head(8)
        bar_chart(latest_shares.index.astype(str), latest_shares.values,
                  f"Leading {label.lower()} commodity shares (FY {last.financial_year})",
                  "Share of reported commodity-table total (%)", f"{share_fig:02d}_top_{prefix}_shares.png", horizontal=True)
        first = df.sort_values("fy_start").iloc[0]
        first_total = numeric(pd.Series([df.loc[first.name, total_col]])).iloc[0] if total_col is not None else numeric(df.loc[first.name, cols]).sum()
        last_total = numeric(pd.Series([df.loc[last.name, total_col]])).iloc[0] if total_col is not None else numeric(df.loc[last.name, cols]).sum()
        start_share = numeric(df.loc[first.name, cols]) / first_total * 100
        end_share = numeric(df.loc[last.name, cols]) / last_total * 100
        delta = (end_share - start_share)[plot_cols].dropna().sort_values(ascending=False).head(8)
        bar_chart(delta.index.astype(str), delta.values,
                  f"Largest increases in {label.lower()} commodity share (FY {first.financial_year} to FY {last.financial_year})",
                  "Change in share (percentage points)", f"{change_fig:02d}_share_change_{prefix}.png", horizontal=True)
        top_names = vals.index[:5]
        plt.figure(figsize=(11, 6))
        for name in top_names:
            plt.plot(df["financial_year"], df[name], marker="o", label=str(name))
        plt.title(f"Selected leading {label.lower()} commodity trends")
        plt.xlabel("Financial year"); plt.ylabel("₹ crore")
        plt.xticks(rotation=45, ha="right"); plt.grid(True, alpha=0.25); plt.legend(fontsize=8)
        plt.tight_layout(); plt.savefig(OUT / f"{trend_fig:02d}_trends_{prefix}_commodities.png", dpi=250); plt.close()

    dr, dcols = direction_table("Direction_Rupees")
    country_cols = [c for c in dcols if re.match(r"^\s*\d+\.", str(c))]
    broad_cols = [c for c in dcols if re.match(r"^\s*(I|II|III|IV|V|VI|VII|VIII)\.", str(c), re.IGNORECASE)]
    total_col = next((c for c in dr.columns if str(c).strip().lower() == "total"), None)
    exports = dr[dr.Flow.eq("export")].sort_values("fy_start")
    imports = dr[dr.Flow.eq("import")].sort_values("fy_start")
    for sub, flow, label in [(exports, "export", "destinations"), (imports, "import", "sources")]:
            latest, first = sub.iloc[-1], sub.iloc[0]
            fy_latest, fy_first = str(latest["Year"]).strip(), str(first["Year"]).strip()
            vals = pd.Series({c: latest[c] for c in country_cols}).dropna().sort_values(ascending=False).head(10)
            top_fig = 25 if flow == "export" else 27
            compare_fig = 26 if flow == "export" else 28
            names = [clean_label(c) for c in vals.index]
            bar_chart(names, vals.values, f"Top ten individual export destination countries (FY {fy_latest})" if flow == "export" else f"Top ten {flow} {label} - individual countries (FY {fy_latest})",
                      "₹ crore", f"{top_fig:02d}_top_{label}.png", horizontal=True)
            compare = pd.DataFrame({f"FY {fy_first}": [first[c] for c in vals.index],
                                    f"FY {fy_latest}": [latest[c] for c in vals.index]}, index=names).fillna(0)
            plt.figure(figsize=(11, 6))
            compare.plot(kind="bar", ax=plt.gca())
            thousands(plt.gca().yaxis)
            plt.title(f"Trade {label} (top ten countries in FY {fy_latest}): FY {fy_first} vs FY {fy_latest}")
            plt.xlabel("Country"); plt.ylabel("₹ crore")
            plt.xticks(rotation=45, ha="right"); plt.grid(axis="y", alpha=0.25); plt.tight_layout()
            plt.savefig(OUT / f"{compare_fig:02d}_partner_endpoint_comparison_{flow}.png", dpi=250); plt.close()

    latest_exp, latest_imp = exports.iloc[-1], imports.iloc[-1]
    fy_latest = str(latest_exp["Year"]).strip()
    balances = pd.Series({clean_label(c): latest_exp[c] - latest_imp[c] for c in country_cols
                              if pd.notna(latest_exp[c]) and pd.notna(latest_imp[c])}).sort_values()
    neg = balances[balances < 0].head(5)
    pos = balances[balances > 0].sort_values(ascending=False).head(5)
    selected = pd.concat([neg, pos]).sort_values()
    bar_chart(selected.index, selected.values,
                  f"Selected bilateral merchandise balances, five largest deficits and surpluses (FY {fy_latest})",
                  "Exports minus imports (₹ crore)", "29_selected_bilateral_balances.png", horizontal=True)

    for j, (row, flow) in enumerate([(latest_exp, "exports"), (latest_imp, "imports")]):
            vals = pd.Series({clean_label(c): row[c] for c in broad_cols}).dropna().sort_values(ascending=True)
            reported_total = numeric(pd.Series([row[total_col]])).iloc[0]
            bar_chart(vals.index, vals.values, f"Broad geographical groups: {flow} (FY {fy_latest})",
                      "₹ crore", f"{30+j}_direction_categories_{flow}.png", horizontal=True)
            plt.figure(figsize=(11, 6.5))
            bars = plt.barh(vals.index, vals.values)
            for bar, v in zip(bars, vals.values):
                plt.text(bar.get_width(), bar.get_y() + bar.get_height() / 2,
                         f"  {v / reported_total * 100:.1f}%", va="center", fontsize=9)
            thousands(plt.gca().xaxis)
            plt.title(f"Broad geographical groups: {flow} (FY {fy_latest}) (₹ crore, with percentage of total)")
            plt.xlabel("₹ crore"); plt.ylabel("Category"); plt.grid(axis="x", alpha=0.25)
            plt.margins(x=0.12); plt.tight_layout()
            plt.savefig(OUT / f"{30+j}_direction_categories_{flow}.png", dpi=250, bbox_inches="tight"); plt.close()

    conc = {}
    for sub, flow in [(exports, "Exports"), (imports, "Imports")]:
            conc[flow] = [
                pd.Series({c: r[c] for c in country_cols}).dropna().sort_values(ascending=False).head(5).sum()
                / numeric(pd.Series([r[total_col]])).iloc[0] * 100
                for _, r in sub.iterrows()
            ]
    fy_labels = [f"{int(y)}-{(int(y) + 1) % 100:02d}" for y in exports.fy_start]
    pd.DataFrame({"financial_year": fy_labels, "exports_top5_share_pct": conc["Exports"],
                      "imports_top5_share_pct": conc["Imports"]}).to_csv(OUT / "top5_partner_concentration.csv", index=False)
    line_chart(fy_labels, [(conc["Exports"], "Exports: top-5 countries"), (conc["Imports"], "Imports: top-5 countries")],
               "Share of total direction-table trade held by the five largest listed country partners, FY 2011–12 to FY 2025–26",
               "Share of reported total (%)", "32_top5_partner_concentration.png")

    print(f"Completed. Output files are saved in: {OUT.resolve()}")
    print("Review chart labels and source-table definitions before including charts in a report.")

if __name__ == "__main__":
    main()


# In[ ]:




