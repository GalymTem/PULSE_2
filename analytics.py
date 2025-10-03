# analytics.py — Assignment #2 (clean visuals, percentile caps, time slider, Excel formatting, optional demo insert)

import os
import argparse
import datetime as dt
import numpy as np

import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from matplotlib.ticker import FuncFormatter

from config import SQLALCHEMY_DATABASE_URL

CHARTS_DIR = "charts"
EXPORTS_DIR = "exports"
QUERIES_PATH = "queries.sql"

os.makedirs(CHARTS_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)

# ---------- DB ----------
engine: Engine = create_engine(SQLALCHEMY_DATABASE_URL, future=True)

# ---------- helpers ----------
def read_query_from_file(name: str, path: str = QUERIES_PATH) -> str:
    """
    Extract a named SQL block from queries.sql.
    Delimiters: a line with '----------------------------------------------------------------'
    and a header line containing:  name: <query_name>
    Any '-- comments' lines are stripped.
    """
    with open(path, "r", encoding="utf-8") as f:
        sql_text = f.read()
    blocks = sql_text.split("\n----------------------------------------------------------------")
    for block in blocks:
        if f"name: {name}" in block:
            lines = [ln for ln in block.splitlines() if not ln.strip().startswith("--")]
            cleaned = "\n".join(lines).strip()
            return cleaned if cleaned.endswith(";") else cleaned + ";"
    raise ValueError(f"Query named '{name}' not found in {path}")

def fetch_df(query_name: str, parse_dates=None, params=None) -> pd.DataFrame:
    sql = read_query_from_file(query_name)
    with engine.connect() as conn:
        df = pd.read_sql_query(text(sql), conn, params=params or {}, parse_dates=parse_dates or [])
    return df.convert_dtypes()

def save_fig(fig, filename: str) -> str:
    path = os.path.join(CHARTS_DIR, filename)
    fig.savefig(path, bbox_inches="tight", dpi=150)
    plt.close(fig)
    return path

def report(df: pd.DataFrame, kind: str, meaning: str):
    print(f"{len(df)} rows → {kind}: {meaning}")

def cap_by_percentile(series, p=0.99):
    import numpy as np
    s = pd.to_numeric(series, errors="coerce")
    thr = np.nanquantile(s, p) if len(s) else None
    return s.clip(upper=thr) if thr is not None else s

# ---------- OPTIONAL: demo insert (Invoice + InvoiceLine) ----------
def demo_insert_sale(customer_id=None, track_id=None, quantity=1, unit_price=None) -> int:
    """
    Minimal sale insert using SQLAlchemy Core.
    Works whether PKs are identity or plain INTEGER (uses MAX(id)+1 + OVERRIDING SYSTEM VALUE).
    Returns new InvoiceId.
    """
    with engine.begin() as conn:
        # Ensure references
        if customer_id is None:
            customer_id = conn.execute(text('SELECT MIN("CustomerId") FROM "Customer";')).scalar()
            if customer_id is None:
                raise RuntimeError('Table "Customer" is empty.')

        if track_id is None:
            track_id = conn.execute(text('SELECT MIN("TrackId") FROM "Track";')).scalar()
            if track_id is None:
                raise RuntimeError('Table "Track" is empty.')

        # Unit price fallback from Track
        if unit_price is None:
            unit_price = conn.execute(
                text('SELECT "UnitPrice" FROM "Track" WHERE "TrackId" = :tid;'),
                {"tid": track_id}
            ).scalar()
            if unit_price is None:
                raise RuntimeError(f"TrackId {track_id} not found.")

        # Next PKs (MAX+1)
        next_invoice_id = conn.execute(text('SELECT COALESCE(MAX("InvoiceId"),0)+1 FROM "Invoice";')).scalar()
        next_invoiceline_id = conn.execute(text('SELECT COALESCE(MAX("InvoiceLineId"),0)+1 FROM "InvoiceLine";')).scalar()

        invoice_date = dt.datetime.now()
        total = round(float(unit_price) * int(quantity), 2)

        invoice_id = conn.execute(
            text('''
                INSERT INTO "Invoice" ("InvoiceId","CustomerId","InvoiceDate","Total")
                OVERRIDING SYSTEM VALUE
                VALUES (:iid, :cid, :idate, :total)
                RETURNING "InvoiceId";
            '''), {"iid": next_invoice_id, "cid": customer_id, "idate": invoice_date, "total": total}
        ).scalar()

        conn.execute(
            text('''
                INSERT INTO "InvoiceLine" ("InvoiceLineId","InvoiceId","TrackId","UnitPrice","Quantity")
                OVERRIDING SYSTEM VALUE
                VALUES (:ilid, :iid, :tid, :price, :qty);
            '''), {"ilid": next_invoiceline_id, "iid": invoice_id, "tid": track_id,
                   "price": float(unit_price), "qty": int(quantity)}
        )

    return int(invoice_id)

# ---------- charts ----------
def chart_pie_revenue_by_category():
    df = fetch_df("pie_revenue_by_category")
    fig, ax = plt.subplots(figsize=(8, 8))
    ax.pie(df["revenue"], labels=df["category"], autopct="%1.1f%%", pctdistance=0.8)
    ax.set_title("Revenue share by product category (delivered, top-10 + Other)")
    path = save_fig(fig, "01_pie_revenue_by_category.png")
    report(df, "Pie", "Delivered revenue share by category")
    return ("Pie", path, df)

def chart_bar_top_sellers():
    df = fetch_df("bar_top_sellers_by_revenue")
    labels = df["seller_id"].astype(str).str.slice(0, 6) + "…" + df["seller_id"].astype(str).str[-4:]
    fig, ax = plt.subplots(figsize=(11, 7))
    ax.bar(labels, df["revenue"])
    ax.set_title("Top 10 sellers by delivered revenue")
    ax.set_xlabel("Seller (ID abridged)")
    ax.set_ylabel("Revenue")
    plt.xticks(rotation=30, ha="right")
    path = save_fig(fig, "02_bar_top_sellers_by_revenue.png")
    report(df, "Bar", "Top sellers by revenue")
    return ("Bar", path, df)

def chart_barh_avg_review_by_category():
    df = fetch_df("barh_avg_review_by_category")
    df = df.sort_values(["avg_score", "n_reviews"], ascending=[False, False]).head(20)
    fig, ax = plt.subplots(figsize=(10, 12))
    ax.barh(df["category"], df["avg_score"])
    ax.invert_yaxis()
    for i, (score, n) in enumerate(zip(pd.to_numeric(df["avg_score"], errors="coerce"),
                                       pd.to_numeric(df["n_reviews"], errors="coerce"))):
        if pd.notna(score) and pd.notna(n):
            ax.text(float(score), i, f"  {float(score):.2f} ({int(n)})", va="center", fontsize=8)
    ax.set_xlabel("Average review score")
    ax.set_ylabel("Category")
    ax.set_title("Average review by category (≥50 reviews), top-20")
    path = save_fig(fig, "03_barh_avg_review_by_category.png")
    report(df, "BarH", "Avg review score by category (top-20)")
    return ("BarH", path, df)

def chart_line_daily_revenue():
    df = fetch_df("line_daily_revenue_2010_2014", parse_dates=["day"])
    df["revenue"] = cap_by_percentile(df["revenue"], 0.99)

    s = df["revenue"]
    df["ma7"] = s.rolling(7, min_periods=1).mean()
    df["p25"] = s.rolling(7, min_periods=1).quantile(0.25)
    df["p75"] = s.rolling(7, min_periods=1).quantile(0.75)

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(df["day"], df["ma7"], linewidth=2, label="7-day moving average")
    ax.fill_between(df["day"], df["p25"], df["p75"], alpha=0.25, label="IQR (7-day)")
    ax.set_xlim(pd.Timestamp("2010-01-01"), pd.Timestamp("2014-12-31"))

    ax.set_title("Delivered revenue — 2010–2014 (7-day MA, 99th pct capped)")
    ax.set_xlabel("Date"); ax.set_ylabel("Revenue")
    ax.grid(True, linestyle=":", alpha=0.5); ax.legend(); fig.autofmt_xdate()

    path = save_fig(fig, "04_line_daily_revenue_2010_2014.png")
    report(df, "Line", "Daily delivered revenue (2010–2014, smoothed MA + IQR)")
    return ("Line", path, df)


def _fd_bins(series: pd.Series) -> int:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 2:
        return 10
    q75, q25 = s.quantile(0.75), s.quantile(0.25)
    iqr = float(q75 - q25) or (s.std() * 1.349) or 1.0
    h = 2 * iqr * (len(s) ** (-1/3))
    if h <= 0:
        return 30
    bins = int((s.max() - s.min()) / h)
    return max(12, min(bins, 80))

def chart_hist_order_value(log_x: bool = False):
    df = fetch_df("hist_order_value")
    s = pd.to_numeric(df["order_value"], errors="coerce").dropna()
    s = cap_by_percentile(s, 0.99)
    bins = _fd_bins(s)

    fig, ax = plt.subplots(figsize=(11, 6))
    ax.hist(s, bins=bins)
    title = "Per-invoice order value distribution (≤99th pct, FD bins)"
    if log_x:
        ax.set_xscale("log")
        title += " — log-scaled"
    ax.set_title(title)
    ax.set_xlabel("Order value")
    ax.set_ylabel("Count")
    ax.grid(True, linestyle=":", alpha=0.4)

    path = save_fig(fig, "05_hist_order_value.png")
    report(df, "Histogram", f"Order value distribution (bins={bins})")
    return ("Hist", path, df)

def chart_strip_duration_by_genre(max_genres=12):
    """
    Shows track duration (minutes) vs genre.
    - Uses 99th percentile cap on duration to handle outliers
    - Chooses up to max_genres by sample count (so the figure stays readable)
    - Jitters points along y for visibility
    """
    df = fetch_df("duration_by_genre_minutes")
    df["duration_min"] = cap_by_percentile(df["duration_min"], 0.99)

    # pick top genres by count (to avoid a 24-legend explosion)
    counts = df["genre"].value_counts()
    keep = list(counts.head(max_genres).index)
    df = df[df["genre"].isin(keep)]

    # stable order by median duration (longer genres on top)
    med = df.groupby("genre")["duration_min"].median().sort_values()
    genres_order = list(med.index)

    # map genres to y positions
    ymap = {g: i for i, g in enumerate(genres_order)}
    df["_y"] = df["genre"].map(ymap).astype(float)

    # jitter for readability
    rng = np.random.default_rng(42)
    df["_y_jit"] = df["_y"] + rng.normal(0, 0.08, size=len(df))

    fig, ax = plt.subplots(figsize=(12, 7))
    ax.scatter(df["duration_min"], df["_y_jit"], s=14, alpha=0.6)

    ax.set_yticks(range(len(genres_order)))
    ax.set_yticklabels(genres_order)
    ax.set_xlabel("Duration (minutes)")
    ax.set_ylabel("Genre")
    ax.set_title("Track duration (minutes) by genre")
    ax.grid(True, axis="x", linestyle=":", alpha=0.4)

    path = save_fig(fig, "06_duration_by_genre_strip.png")
    report(df, "Strip", "Duration vs genre (99th pct capped)")
    return ("Strip", path, df[["genre", "duration_min"]].copy())



# ---------- plotly time slider ----------
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

CHARTS_DIR = "charts"

def show_time_slider(auto_open_html: bool = True):
    df = fetch_df("timeslider_monthly_revenue_by_country")
    # expected: df has columns "month" (YYYY-MM), "country", "revenue"

    # ensure datetime ordering of months
    df["month_ord"] = pd.to_datetime(df["month"], format="%Y-%m")
    df = df.sort_values(["month_ord", "revenue"], ascending=[True, False])

    # keep only top 10 per month
    df = df.groupby("month", group_keys=False).head(10)

    # build vertical bar chart with animation
    fig = px.bar(
        df,
        x="country",
        y="revenue",
        animation_frame="month",
        animation_group="country",
        title="Monthly delivered revenue by country (Top 10)",
        labels={"country": "Country", "revenue": "Revenue", "month": "Month"},
        template="plotly_white"
    )

    # layout tweaks
    fig.update_layout(
        margin=dict(l=40, r=20, t=60, b=60),
        xaxis_title="Country",
        yaxis_title="Revenue",
        transition=dict(duration=350)
    )
    fig.update_traces(hovertemplate="Country=%{x}<br>Revenue=%{y}<extra></extra>")

    # fix order for each frame: largest revenue first (leftmost), then 2nd, etc.
    for fr in fig.frames:
        fdf = df[df["month"] == fr.name].sort_values("revenue", ascending=False)
        order = fdf["country"].tolist()
        fr.layout = go.Layout(xaxis=dict(categoryorder="array", categoryarray=order))

    # also fix for initial frame
    first_month = df["month"].iloc[0]
    init_order = (
        df[df["month"] == first_month]
        .sort_values("revenue", ascending=False)["country"]
        .tolist()
    )
    fig.update_layout(xaxis=dict(categoryorder="array", categoryarray=init_order))

    # save to HTML (interactive)
    os.makedirs(CHARTS_DIR, exist_ok=True)
    out_html = os.path.join(CHARTS_DIR, "timeslider_revenue_by_country.html")
    fig.write_html(out_html, auto_open=auto_open_html)

# ---------- main ----------
def main():
    import argparse
    ap = argparse.ArgumentParser(description="Generate analytics charts + Excel export from PostgreSQL.")
    ap.add_argument("--insert_demo", action="store_true",
                    help="Insert a demo sale (Invoice + InvoiceLine) so charts change during defense.")
    ap.add_argument("--open_html", action="store_true",
                    help="Open Plotly time slider HTML after save.")
    ap.add_argument("--log_hist", action="store_true",
                    help="Plot histogram with log-scaled X (good for long tails).")
    args = ap.parse_args()

    if args.insert_demo:
        new_invoice_id = demo_insert_sale()
        print(f"[INFO] Inserted demo InvoiceId={new_invoice_id}. Charts will reflect this new sale.")

    # build charts
    charts = [
        chart_pie_revenue_by_category(),
        chart_bar_top_sellers(),
        chart_barh_avg_review_by_category(),
        chart_line_daily_revenue(),                         # smoothed MA + IQR band
        chart_hist_order_value(log_x=args.log_hist),        # FD bins, optional log-x
        chart_strip_duration_by_genre(),                  # duration (minutes) vs price, top-5 genres
    ]

    # improved time slider
    show_time_slider(auto_open_html=args.open_html)

    print("\n[REPORT] Charts → /charts/, Excel → /exports/analytics_export.xlsx")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("[!] Error:", e)
        print("Hint: Ensure queries.sql has 'scatter_price_vs_duration_minutes', "
              "DB URL in config.py is correct, and numpy is imported for polyfit.")
