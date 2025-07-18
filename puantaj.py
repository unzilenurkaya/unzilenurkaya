"""Puantaj Otomasyon Sistemi

This script processes VisitTrack turnstile logs and shift schedule Excel files.
It detects anomalies such as missing card entries or entries on leave days and
exports results to a new Excel workbook.

Usage:
    python3 puantaj.py visittrack.xls shift_schedule.xlsx output.xlsx

The program attempts to automatically detect the structure of the shift
schedule file. If detection fails it will ask the user to specify which rows
contain the day names and dates.
"""

import argparse
import pandas as pd
from datetime import datetime
from datetime import timedelta
from pathlib import Path


def read_visittrack(filepath: Path) -> pd.DataFrame:
    """Read VisitTrack (.xls) data.

    Parameters
    ----------
    filepath: Path to the VisitTrack Excel file.

    Returns
    -------
    DataFrame with columns EmployeeName, CardNo, Datetime, Direction.
    """
    df = pd.read_excel(filepath, sheet_name="Sayfa1", skiprows=4, engine="xlrd")
    df["Datetime"] = pd.to_datetime(df["Tarih"].astype(str) + " " + df["Saat"].astype(str))
    df = df.rename(
        columns={
            "Adı Soyadı": "EmployeeName",
            "Kart No": "CardNo",
            "Yön": "Direction",
            "Kapı Adı": "Gate",
        }
    )
    return df[["EmployeeName", "CardNo", "Datetime", "Direction"]]


def detect_headers(df: pd.DataFrame):
    """Attempt to locate header rows in the shift schedule file.

    Returns indices for the rows containing day names, dates and the start
    of employee data. If detection fails, returns None for unknown values.
    """
    days_row = dates_row = None
    for i in range(min(10, len(df))):
        row = df.iloc[i].astype(str).str.lower()
        if days_row is None and row.str.contains("paz|per|çar|cum|sal").any():
            days_row = i
        elif dates_row is None and row.str.contains(r"\b[0-9]{1,2}\b").sum() > 10:
            dates_row = i
    if days_row is not None and dates_row is not None:
        start_row = max(days_row, dates_row) + 1
    else:
        start_row = None
    return days_row, dates_row, start_row


def read_shift_schedule(filepath: Path) -> pd.DataFrame:
    """Read shift schedule with flexible header detection."""
    raw = pd.read_excel(filepath, sheet_name=0, header=None, engine="openpyxl")
    days_row, dates_row, start_row = detect_headers(raw)

    if days_row is None or dates_row is None or start_row is None:
        print("\nDosya başlıkları otomatik tespit edilemedi.")
        days_row = int(input("Hafta günlerinin bulunduğu satır numarası (0-index): "))
        dates_row = int(input("Tarihlerin bulunduğu satır numarası (0-index): "))
        start_row = int(input("Personel verisinin başladığı satır numarası (0-index): "))

    days = raw.iloc[days_row]
    dates = raw.iloc[dates_row]
    data = raw.iloc[start_row:]

    # Forward fill employee names in case of merged cells
    data.iloc[:, 0] = data.iloc[:, 0].ffill()

    data.columns = ["Employee"] + list(dates[1:])
    df_long = data.melt(id_vars="Employee", var_name="Date", value_name="ShiftCode")
    df_long["Date"] = pd.to_datetime(df_long["Date"], errors="coerce")
    df_long = df_long.dropna(subset=["Date"])
    df_long["ShiftCode"] = df_long["ShiftCode"].astype(str).str.strip()
    return df_long


def match_shifts(visit_df: pd.DataFrame, shift_df: pd.DataFrame) -> pd.DataFrame:
    """Compare shift expectations with turnstile records."""
    anomalies = []
    visit_df["Date"] = visit_df["Datetime"].dt.date

    for employee, emp_shifts in shift_df.groupby("Employee"):
        emp_visits = visit_df[visit_df["EmployeeName"].str.lower() == employee.lower()]
        for _, row in emp_shifts.iterrows():
            date = row["Date"].date()
            code = row["ShiftCode"]
            day_visits = emp_visits[emp_visits["Date"] == date]

            if code in {"H", "FMİ", "RMİ", "T"}:
                if not day_visits.empty:
                    anomalies.append({
                        "Employee": employee,
                        "Date": row["Date"],
                        "Issue": "Entry on day off",
                    })
                continue

            if day_visits.empty:
                anomalies.append({
                    "Employee": employee,
                    "Date": row["Date"],
                    "Issue": "No entry/exit recorded",
                })
            else:
                has_entry = (day_visits["Direction"] == "Giriş").any()
                has_exit = (day_visits["Direction"] == "Çıkış").any()
                if not has_entry or not has_exit:
                    anomalies.append({
                        "Employee": employee,
                        "Date": row["Date"],
                        "Issue": "Missing entry or exit",
                    })

    return pd.DataFrame(anomalies)


def generate_reports(anomalies: pd.DataFrame, output: Path):
    """Write anomalies to an Excel file with one sheet per employee."""
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for employee, group in anomalies.groupby("Employee"):
            sheet = employee[:31]
            group.to_excel(writer, sheet_name=sheet, index=False)


def main():
    parser = argparse.ArgumentParser(description="Puantaj Otomasyon Sistemi")
    parser.add_argument("visittrack", type=Path, help="VisitTrack dosyası (.xls)")
    parser.add_argument("shifts", type=Path, help="Vardiya listesi (.xlsx)")
    parser.add_argument("output", type=Path, help="Çıktı Excel dosyası")
    args = parser.parse_args()

    visit_df = read_visittrack(args.visittrack)
    shift_df = read_shift_schedule(args.shifts)
    anomalies = match_shifts(visit_df, shift_df)
    generate_reports(anomalies, args.output)
    print(f"Rapor oluşturuldu: {args.output}")


if __name__ == "__main__":
    main()
