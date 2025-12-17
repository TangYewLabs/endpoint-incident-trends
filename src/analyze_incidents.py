import argparse
from pathlib import Path
import pandas as pd


def safe_int(x):
    try:
        if pd.isna(x):
            return None
        return int(float(x))
    except Exception:
        return None


def main():
    parser = argparse.ArgumentParser(
        description="Endpoint Incident Trend Analyzer: CSV -> Excel report"
    )
    parser.add_argument("csv_path", help="Path to incident CSV export")
    parser.add_argument(
        "--out",
        default="reports/incident_trends_report.xlsx",
        help="Output Excel report path",
    )
    args = parser.parse_args()

    csv_path = Path(args.csv_path)
    out_path = Path(args.out)

    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(csv_path)

    required_cols = [
        "incident_id",
        "opened_at",
        "user_role",
        "device_type",
        "issue_category",
        "priority",
        "resolution_minutes",
        "resolved",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Normalize
    df["opened_at"] = pd.to_datetime(df["opened_at"], errors="coerce")
    df["resolved"] = df["resolved"].astype(str).str.strip().str.lower().map(
        {"yes": True, "no": False}
    )
    df["resolution_minutes"] = df["resolution_minutes"].apply(safe_int)
    df["is_executive"] = df["user_role"].astype(str).str.strip().str.lower().eq("executive")

    # Basic metrics
    total = len(df)
    resolved_count = int(df["resolved"].fillna(False).sum())
    unresolved_count = total - resolved_count

    # Frequency of issues
    freq = (
        df.groupby("issue_category", dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
    )
    freq["percent"] = (freq["count"] / max(total, 1) * 100).round(1)

    # Time metrics by issue
    time_by_issue = (
        df[df["resolution_minutes"].notna()]
        .groupby("issue_category")["resolution_minutes"]
        .agg(["count", "mean", "median", "max"])
        .reset_index()
        .rename(columns={"count": "n_with_time", "mean": "avg_minutes"})
    )
    for col in ["avg_minutes", "median", "max"]:
        time_by_issue[col] = time_by_issue[col].round(1)

    # Executive impact
    exec_view = (
        df.groupby(["is_executive", "issue_category"])
        .size()
        .reset_index(name="count")
        .sort_values(["is_executive", "count"], ascending=[False, False])
    )
    exec_view["user_group"] = exec_view["is_executive"].map({True: "Executive", False: "Non-Executive"})
    exec_view = exec_view[["user_group", "issue_category", "count"]]

    # SLA risk indicator (simple heuristic)
    sla_thresholds = {"P1": 60, "P2": 240, "P3": 1440, "P4": 2880}
    df["sla_minutes"] = df["priority"].astype(str).map(sla_thresholds)
    df["sla_breached"] = (
        df["resolution_minutes"].notna()
        & df["sla_minutes"].notna()
        & (df["resolution_minutes"] > df["sla_minutes"])
    )

    sla_summary = (
        df[df["resolution_minutes"].notna() & df["sla_minutes"].notna()]
        .groupby("priority")["sla_breached"]
        .agg(total_with_sla="count", breaches="sum")
        .reset_index()
    )
    sla_summary["breach_rate_percent"] = (
        (sla_summary["breaches"] / sla_summary["total_with_sla"] * 100).round(1)
    )

    # Recommendations (simple rule-based)
    top3 = freq.head(3)["issue_category"].tolist()
    recs = []
    for issue in top3:
        if "vpn" in issue.lower():
            recs.append("VPN: validate MFA/token health, certificate lifecycle, and client version standardization.")
        elif "teams" in issue.lower():
            recs.append("Teams: standardize audio/video device profiles and confirm driver/firmware baselines.")
        elif "wi-fi" in issue.lower() or "wifi" in issue.lower():
            recs.append("Wi-Fi: enforce known-good driver versions and check access point roaming/coverage hotspots.")
        elif "o365" in issue.lower():
            recs.append("O365: review conditional access/MFA failure modes and sign-in logs for recurring patterns.")
        elif "conference" in issue.lower() or "av" in issue.lower():
            recs.append("Conference Room AV: create a pre-meeting health checklist and standardize room profiles.")
        else:
            recs.append(f"{issue}: review recurring causes and define a standard fix playbook.")

    summary_rows = [
        ("Total incidents", total),
        ("Resolved incidents", resolved_count),
        ("Unresolved incidents", unresolved_count),
        ("Top issue category", top3[0] if top3 else "N/A"),
    ]
    summary = pd.DataFrame(summary_rows, columns=["metric", "value"])
    recommendations = pd.DataFrame({"recommendation": recs})

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        freq.to_excel(writer, index=False, sheet_name="Top_Issues")
        time_by_issue.to_excel(writer, index=False, sheet_name="Time_By_Issue")
        exec_view.to_excel(writer, index=False, sheet_name="Executive_Impact")
        sla_summary.to_excel(writer, index=False, sheet_name="SLA_Risk")
        recommendations.to_excel(writer, index=False, sheet_name="Recommendations")
        df.to_excel(writer, index=False, sheet_name="Raw_Data")

    print("Endpoint Incident Trend Analyzer")
    print(f"Input:  {csv_path}")
    print(f"Output: {out_path}")
    print("\nTop Issues:")
    print(freq.head(5).to_string(index=False))
    print("\nRecommendations:")
    for r in recs:
        print(f"- {r}")


if __name__ == "__main__":
    main()
