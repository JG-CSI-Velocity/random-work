"""Count ILSFILE2 debit-decline rows per day and per month, output to Excel."""

from pathlib import Path
import re
import pandas as pd

BASE = Path(r"C:\Users\james.gilmore\Downloads\TNB_VelocityArchive_Jan-Mar_2026")
OUT = BASE / "ILSFILE2_decline_counts.xlsx"
COLUMNS = ["Date", "Acct#", "TranType", "Amt", "Last4Card", "Description", "TerminalID"]
EXPECTED_FIELDS = len(COLUMNS)
DATE_DIR = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def count_file(path: Path) -> tuple[int, int]:
    """Return (valid_rows, malformed_rows). Blank lines are skipped."""
    valid = malformed = 0
    with path.open("r", encoding="utf-8", errors="replace", newline="") as fh:
        for line in fh:
            stripped = line.strip()
            if not stripped:
                continue
            if stripped.count(",") == EXPECTED_FIELDS - 1:
                valid += 1
            else:
                malformed += 1
    return valid, malformed


def scan() -> pd.DataFrame:
    rows = []
    for folder in sorted(p for p in BASE.iterdir() if p.is_dir() and DATE_DIR.match(p.name)):
        target = folder / "ILSFILE2"
        if not target.exists():
            rows.append({"date": folder.name, "count": 0, "malformed": 0,
                         "status": "missing_file", "note": "ILSFILE2 not present"})
            continue
        valid, malformed = count_file(target)
        status = "empty" if valid == 0 and malformed == 0 else "ok"
        rows.append({"date": folder.name, "count": valid, "malformed": malformed,
                     "status": status, "note": ""})
    return pd.DataFrame(rows)


def main() -> None:
    if not BASE.exists():
        raise SystemExit(f"Base path not found: {BASE}")

    daily = scan()
    if daily.empty:
        raise SystemExit("No daily folders found.")

    daily["date"] = pd.to_datetime(daily["date"])
    daily = daily.sort_values("date").reset_index(drop=True)
    daily["month"] = daily["date"].dt.strftime("%Y-%m")

    monthly = (
        daily.groupby("month")
        .agg(declines=("count", "sum"),
             days_with_data=("status", lambda s: (s == "ok").sum()),
             empty_days=("status", lambda s: (s == "empty").sum()),
             missing_files=("status", lambda s: (s == "missing_file").sum()),
             malformed_rows=("malformed", "sum"))
        .reset_index()
    )

    total = pd.DataFrame([{
        "month": "TOTAL",
        "declines": daily["count"].sum(),
        "days_with_data": (daily["status"] == "ok").sum(),
        "empty_days": (daily["status"] == "empty").sum(),
        "missing_files": (daily["status"] == "missing_file").sum(),
        "malformed_rows": daily["malformed"].sum(),
    }])
    monthly = pd.concat([monthly, total], ignore_index=True)

    daily_out = daily[["date", "count", "status", "malformed", "note"]].copy()
    daily_out["date"] = daily_out["date"].dt.strftime("%Y-%m-%d")

    with pd.ExcelWriter(OUT, engine="openpyxl") as xl:
        daily_out.to_excel(xl, sheet_name="Daily", index=False)
        monthly.to_excel(xl, sheet_name="Monthly", index=False)

    print(f"Wrote {OUT}")
    print(f"Days scanned: {len(daily)}  |  Total declines: {daily['count'].sum()}")


if __name__ == "__main__":
    main()
