from __future__ import annotations

import json
from pathlib import Path

import duckdb
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np
import pandas as pd


CSV_PATH = Path("medicaid-provider-spending.csv")
OUT_DIR = Path("eda_outputs")
OUT_DIR.mkdir(exist_ok=True)

SAMPLE_N = 100_000


def compact_number(x: float, _pos: int | None = None) -> str:
    abs_x = abs(x)
    if abs_x >= 1_000_000_000:
        return f"{x/1_000_000_000:.1f}B"
    if abs_x >= 1_000_000:
        return f"{x/1_000_000:.1f}M"
    if abs_x >= 1_000:
        return f"{x/1_000:.1f}K"
    return f"{x:.0f}"


def compact_currency(x: float, _pos: int | None = None) -> str:
    abs_x = abs(x)
    sign = "-" if x < 0 else ""
    if abs_x >= 1_000_000_000:
        return f"{sign}${abs_x/1_000_000_000:.1f}B"
    if abs_x >= 1_000_000:
        return f"{sign}${abs_x/1_000_000:.1f}M"
    if abs_x >= 1_000:
        return f"{sign}${abs_x/1_000:.1f}K"
    return f"{sign}${abs_x:.0f}"


def apply_plot_style() -> None:
    plt.style.use("seaborn-v0_8-whitegrid")
    plt.rcParams.update(
        {
            "figure.dpi": 160,
            "axes.titlesize": 14,
            "axes.labelsize": 11,
            "xtick.labelsize": 10,
            "ytick.labelsize": 10,
            "legend.fontsize": 10,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "grid.alpha": 0.3,
        }
    )


def main() -> None:
    if not CSV_PATH.exists():
        raise FileNotFoundError(f"Missing input file: {CSV_PATH}")

    apply_plot_style()

    con = duckdb.connect(database=":memory:")

    # Keep ingestion explicit so types are predictable for downstream summaries/plots.
    csv_scan = f"""
        read_csv_auto(
            '{CSV_PATH.as_posix()}',
            header=true,
            types={{
                'BILLING_PROVIDER_NPI_NUM': 'VARCHAR',
                'SERVICING_PROVIDER_NPI_NUM': 'VARCHAR',
                'HCPCS_CODE': 'VARCHAR',
                'CLAIM_FROM_MONTH': 'VARCHAR',
                'TOTAL_UNIQUE_BENEFICIARIES': 'DOUBLE',
                'TOTAL_CLAIMS': 'DOUBLE',
                'TOTAL_PAID': 'DOUBLE'
            }}
        )
    """
    con.execute(f"CREATE VIEW data AS SELECT * FROM {csv_scan}")

    preview_df = con.sql("SELECT * FROM data LIMIT 5").df()

    report_row = con.sql(
        """
        SELECT
            COUNT(*) AS total_rows,
            COUNT(DISTINCT CLAIM_FROM_MONTH) AS unique_claim_months,
            MIN(CLAIM_FROM_MONTH) AS claim_month_min,
            MAX(CLAIM_FROM_MONTH) AS claim_month_max,
            COUNT(DISTINCT HCPCS_CODE) AS unique_hcpcs_codes,
            SUM(CASE WHEN BILLING_PROVIDER_NPI_NUM IS NULL THEN 1 ELSE 0 END) AS missing_BILLING_PROVIDER_NPI_NUM,
            SUM(CASE WHEN SERVICING_PROVIDER_NPI_NUM IS NULL THEN 1 ELSE 0 END) AS missing_SERVICING_PROVIDER_NPI_NUM,
            SUM(CASE WHEN HCPCS_CODE IS NULL THEN 1 ELSE 0 END) AS missing_HCPCS_CODE,
            SUM(CASE WHEN CLAIM_FROM_MONTH IS NULL THEN 1 ELSE 0 END) AS missing_CLAIM_FROM_MONTH,
            SUM(CASE WHEN TOTAL_UNIQUE_BENEFICIARIES IS NULL THEN 1 ELSE 0 END) AS missing_TOTAL_UNIQUE_BENEFICIARIES,
            SUM(CASE WHEN TOTAL_CLAIMS IS NULL THEN 1 ELSE 0 END) AS missing_TOTAL_CLAIMS,
            SUM(CASE WHEN TOTAL_PAID IS NULL THEN 1 ELSE 0 END) AS missing_TOTAL_PAID
        FROM data
        """
    ).df().iloc[0]

    numeric_summary = con.sql(
        """
        SELECT * FROM (
            SELECT
                'TOTAL_UNIQUE_BENEFICIARIES' AS column,
                COUNT(TOTAL_UNIQUE_BENEFICIARIES) AS count,
                AVG(TOTAL_UNIQUE_BENEFICIARIES) AS mean,
                STDDEV_SAMP(TOTAL_UNIQUE_BENEFICIARIES) AS std,
                MIN(TOTAL_UNIQUE_BENEFICIARIES) AS min,
                MAX(TOTAL_UNIQUE_BENEFICIARIES) AS max,
                SUM(CASE WHEN TOTAL_UNIQUE_BENEFICIARIES = 0 THEN 1 ELSE 0 END) AS zeros,
                SUM(CASE WHEN TOTAL_UNIQUE_BENEFICIARIES < 0 THEN 1 ELSE 0 END) AS negatives
            FROM data
            UNION ALL
            SELECT
                'TOTAL_CLAIMS' AS column,
                COUNT(TOTAL_CLAIMS) AS count,
                AVG(TOTAL_CLAIMS) AS mean,
                STDDEV_SAMP(TOTAL_CLAIMS) AS std,
                MIN(TOTAL_CLAIMS) AS min,
                MAX(TOTAL_CLAIMS) AS max,
                SUM(CASE WHEN TOTAL_CLAIMS = 0 THEN 1 ELSE 0 END) AS zeros,
                SUM(CASE WHEN TOTAL_CLAIMS < 0 THEN 1 ELSE 0 END) AS negatives
            FROM data
            UNION ALL
            SELECT
                'TOTAL_PAID' AS column,
                COUNT(TOTAL_PAID) AS count,
                AVG(TOTAL_PAID) AS mean,
                STDDEV_SAMP(TOTAL_PAID) AS std,
                MIN(TOTAL_PAID) AS min,
                MAX(TOTAL_PAID) AS max,
                SUM(CASE WHEN TOTAL_PAID = 0 THEN 1 ELSE 0 END) AS zeros,
                SUM(CASE WHEN TOTAL_PAID < 0 THEN 1 ELSE 0 END) AS negatives
            FROM data
        )
        """
    ).df()

    monthly_df = con.sql(
        """
        SELECT
            CLAIM_FROM_MONTH,
            COUNT(*) AS rows,
            SUM(TOTAL_UNIQUE_BENEFICIARIES) AS TOTAL_UNIQUE_BENEFICIARIES_sum,
            SUM(TOTAL_CLAIMS) AS TOTAL_CLAIMS_sum,
            SUM(TOTAL_PAID) AS TOTAL_PAID_sum
        FROM data
        GROUP BY CLAIM_FROM_MONTH
        ORDER BY CLAIM_FROM_MONTH
        """
    ).df()

    hcpcs_df = con.sql(
        """
        SELECT
            HCPCS_CODE,
            COUNT(*) AS rows,
            SUM(TOTAL_UNIQUE_BENEFICIARIES) AS TOTAL_UNIQUE_BENEFICIARIES_sum,
            SUM(TOTAL_CLAIMS) AS TOTAL_CLAIMS_sum,
            SUM(TOTAL_PAID) AS TOTAL_PAID_sum
        FROM data
        GROUP BY HCPCS_CODE
        ORDER BY TOTAL_PAID_sum DESC NULLS LAST
        """
    ).df()

    # Sample rows for quantiles and row-level plots.
    try:
        sample_df = con.sql(
            f"""
            SELECT CLAIM_FROM_MONTH, HCPCS_CODE, TOTAL_UNIQUE_BENEFICIARIES, TOTAL_CLAIMS, TOTAL_PAID
            FROM data
            USING SAMPLE reservoir({SAMPLE_N} ROWS)
            """
        ).df()
    except Exception:
        # Fallback if SAMPLE syntax differs by DuckDB build.
        sample_df = con.sql(
            f"""
            SELECT CLAIM_FROM_MONTH, HCPCS_CODE, TOTAL_UNIQUE_BENEFICIARIES, TOTAL_CLAIMS, TOTAL_PAID
            FROM data
            WHERE random() < 0.002
            LIMIT {SAMPLE_N}
            """
        ).df()

    if sample_df.empty:
        sample_quantiles = pd.DataFrame()
    else:
        sample_quantiles = (
            sample_df[["TOTAL_UNIQUE_BENEFICIARIES", "TOTAL_CLAIMS", "TOTAL_PAID"]]
            .quantile([0.01, 0.05, 0.25, 0.50, 0.75, 0.95, 0.99])
            .transpose()
            .reset_index()
            .rename(columns={"index": "column"})
        )

    top_hcpcs = hcpcs_df.head(15).copy()

    # Save tabular outputs
    preview_df.to_csv(OUT_DIR / "data_preview_head.csv", index=False)
    numeric_summary.to_csv(OUT_DIR / "numeric_summary_exact_moments.csv", index=False)
    sample_quantiles.to_csv(OUT_DIR / "numeric_quantiles_sample.csv", index=False)
    monthly_df.to_csv(OUT_DIR / "monthly_summary.csv", index=False)
    hcpcs_df.head(200).to_csv(OUT_DIR / "top_hcpcs_summary_top200.csv", index=False)
    sample_df.to_csv(OUT_DIR / "eda_sample_rows.csv", index=False)

    # Plots
    plot_monthly = monthly_df[monthly_df["CLAIM_FROM_MONTH"].notna()].copy()
    if not plot_monthly.empty:
        plot_monthly["CLAIM_FROM_MONTH"] = pd.to_datetime(plot_monthly["CLAIM_FROM_MONTH"], format="%Y-%m")
        plot_monthly = plot_monthly.sort_values("CLAIM_FROM_MONTH")
        month_labels = plot_monthly["CLAIM_FROM_MONTH"].dt.strftime("%Y-%m")
        tick_step = max(len(plot_monthly) // 12, 1)
        tick_positions = np.arange(0, len(plot_monthly), tick_step)

        plt.figure(figsize=(13, 6))
        plt.plot(
            month_labels,
            plot_monthly["TOTAL_PAID_sum"],
            linewidth=2.4,
            color="#005f73",
        )
        plt.xticks(tick_positions, month_labels.iloc[tick_positions], rotation=45, ha="right")
        plt.gca().yaxis.set_major_formatter(FuncFormatter(compact_currency))
        plt.ylabel("Total Paid ($)")
        plt.xlabel("Claim Month")
        plt.title("Total Paid by Claim Month")
        plt.tight_layout()
        plt.savefig(OUT_DIR / "monthly_total_paid.png", dpi=220)
        plt.close()

        plt.figure(figsize=(13, 6))
        plt.plot(
            month_labels,
            plot_monthly["rows"],
            linewidth=2.4,
            color="#0a9396",
        )
        plt.xticks(tick_positions, month_labels.iloc[tick_positions], rotation=45, ha="right")
        plt.gca().yaxis.set_major_formatter(FuncFormatter(compact_number))
        plt.ylabel("Row Count")
        plt.xlabel("Claim Month")
        plt.title("Rows by Claim Month")
        plt.tight_layout()
        plt.savefig(OUT_DIR / "monthly_row_count.png", dpi=220)
        plt.close()

    if not top_hcpcs.empty:
        plot_h = top_hcpcs.sort_values("TOTAL_PAID_sum", ascending=True)
        plt.figure(figsize=(11, 8))
        bars = plt.barh(plot_h["HCPCS_CODE"].astype(str), plot_h["TOTAL_PAID_sum"], color="#94d2bd")
        plt.gca().xaxis.set_major_formatter(FuncFormatter(compact_currency))
        plt.xlabel("Total Paid ($)")
        plt.ylabel("HCPCS Code")
        plt.title("Top 15 HCPCS Codes by Total Paid")
        x_max = plot_h["TOTAL_PAID_sum"].max()
        for bar in bars:
            w = bar.get_width()
            plt.text(w + x_max * 0.01, bar.get_y() + bar.get_height() / 2, compact_currency(w), va="center", fontsize=9)
        plt.tight_layout()
        plt.savefig(OUT_DIR / "top_hcpcs_total_paid.png", dpi=220)
        plt.close()

    if not sample_df.empty:
        paid = sample_df["TOTAL_PAID"].dropna()
        if not paid.empty:
            plt.figure(figsize=(10, 6))
            plt.hist(np.log1p(paid.clip(lower=0)), bins=70, color="#ee9b00", alpha=0.9, edgecolor="white", linewidth=0.3)
            plt.xlabel("log1p(TOTAL_PAID)")
            plt.ylabel("Frequency")
            plt.title("Distribution of Per-Row Total Paid (Sample)")
            plt.tight_layout()
            plt.savefig(OUT_DIR / "dist_total_paid_log_sample.png", dpi=220)
            plt.close()

        scatter_df = sample_df[["TOTAL_CLAIMS", "TOTAL_PAID"]].dropna()
        if not scatter_df.empty:
            scatter_df = scatter_df.sample(n=min(20_000, len(scatter_df)), random_state=42)
            plt.figure(figsize=(8, 7))
            plt.scatter(
                np.log1p(scatter_df["TOTAL_CLAIMS"].clip(lower=0)),
                np.log1p(scatter_df["TOTAL_PAID"].clip(lower=0)),
                s=9,
                alpha=0.22,
                color="#bb3e03",
                linewidths=0,
            )
            plt.xlabel("log1p(TOTAL_CLAIMS)")
            plt.ylabel("log1p(TOTAL_PAID)")
            plt.title("Claims vs Paid per Row (Sample)")
            plt.tight_layout()
            plt.savefig(OUT_DIR / "scatter_claims_vs_paid_log_sample.png", dpi=220)
            plt.close()

    missing_counts = {
        "BILLING_PROVIDER_NPI_NUM": int(report_row["missing_BILLING_PROVIDER_NPI_NUM"]),
        "SERVICING_PROVIDER_NPI_NUM": int(report_row["missing_SERVICING_PROVIDER_NPI_NUM"]),
        "HCPCS_CODE": int(report_row["missing_HCPCS_CODE"]),
        "CLAIM_FROM_MONTH": int(report_row["missing_CLAIM_FROM_MONTH"]),
        "TOTAL_UNIQUE_BENEFICIARIES": int(report_row["missing_TOTAL_UNIQUE_BENEFICIARIES"]),
        "TOTAL_CLAIMS": int(report_row["missing_TOTAL_CLAIMS"]),
        "TOTAL_PAID": int(report_row["missing_TOTAL_PAID"]),
    }

    report = {
        "input_file": str(CSV_PATH),
        "total_rows": int(report_row["total_rows"]),
        "columns": list(preview_df.columns),
        "missing_counts": missing_counts,
        "claim_month_min": report_row["claim_month_min"],
        "claim_month_max": report_row["claim_month_max"],
        "unique_claim_months": int(report_row["unique_claim_months"]),
        "unique_hcpcs_codes": int(report_row["unique_hcpcs_codes"]),
        "sample_rows_for_eda": int(len(sample_df)),
    }
    (OUT_DIR / "eda_report.json").write_text(json.dumps(report, indent=2))

    with (OUT_DIR / "eda_report.txt").open("w") as f:
        f.write("Medicaid Provider Spending EDA Summary\n")
        f.write("=" * 40 + "\n\n")
        for k, v in report.items():
            f.write(f"{k}: {v}\n")
        f.write("\nExact-moment numeric summary (mean/std/min/max):\n")
        f.write(numeric_summary.to_string(index=False))
        f.write("\n\nSample-based quantiles:\n")
        f.write(sample_quantiles.to_string(index=False) if not sample_quantiles.empty else "No sample rows.\n")
        f.write("\n\nTop 15 HCPCS by total paid:\n")
        f.write(top_hcpcs.to_string(index=False) if not top_hcpcs.empty else "No HCPCS data.\n")

    print(f"Done. Outputs saved to {OUT_DIR.resolve()}")
    print(f"Rows processed: {report['total_rows']:,}")
    print("Files:")
    for p in sorted(OUT_DIR.iterdir()):
        print(f" - {p.name}")


if __name__ == "__main__":
    main()
