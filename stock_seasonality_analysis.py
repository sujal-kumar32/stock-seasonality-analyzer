import tkinter as tk
from tkinter import messagebox
import yfinance as yf
import pandas as pd
from calendar import month_abbr
from datetime import datetime
import os

# ---------------- CONFIG ----------------
LOOKBACK_YEARS = 15
MONTH_NAMES = list(month_abbr)[1:]  # Jan–Dec

POPULAR_TICKERS = {
    "US Stocks": [
        "AAPL (Apple)", "MSFT (Microsoft)", "TSLA (Tesla)",
        "GOOGL (Google)", "AMZN (Amazon)", "META (Meta)",
        "NFLX (Netflix)", "NVDA (NVIDIA)"
    ],
    "Indian Stocks (NSE)": [
        "RELIANCE.NS", "TCS.NS", "INFY.NS",
        "HDFCBANK.NS", "ICICIBANK.NS", "SBIN.NS"
    ]
}

# ------------- ANALYSIS LOGIC -------------
def analyze_all_months_performance(tickers, years):
    end_date = pd.to_datetime("today").normalize()
    start_date = end_date - pd.DateOffset(years=years)

    all_results = []

    for ticker in tickers:
        data = yf.download(ticker, start=start_date, end=end_date, progress=False)

        if data.empty:
            continue

        ticker_results = {"Ticker": ticker}

        for month_num, month_name in enumerate(MONTH_NAMES, start=1):
            monthly_returns = []

            for year in range(start_date.year, end_date.year + 1):
                month_data = data[
                    (data.index.month == month_num) &
                    (data.index.year == year)
                ]

                if month_data.empty or len(month_data) < 5:
                    continue

                start_price = float(month_data["Open"].iloc[0])
                end_price = float(month_data["Close"].iloc[-1])

                monthly_returns.append((end_price - start_price) / start_price)

            if monthly_returns:
                series = pd.Series(monthly_returns)
                avg_return = float(series.mean()) * 100
                hit_rate = (series.gt(0).sum() / len(series)) * 100
            else:
                avg_return = 0.0
                hit_rate = 0.0

            ticker_results[f"{month_name} Avg. Rtn (%)"] = avg_return
            ticker_results[f"{month_name} Hit Rate (%)"] = hit_rate

        all_results.append(ticker_results)

    return pd.DataFrame(all_results)

# ---------------- GUI LOGIC ----------------
def run_analysis():
    raw_input = stock_entry.get()

    if not raw_input.strip():
        messagebox.showwarning("Input Error", "Please enter at least one stock ticker.")
        return

    tickers = [t.strip().upper() for t in raw_input.split(",") if t.strip()]

    status_label.config(text="Running analysis... Please wait ⏳", fg="orange")
    root.update()

    try:
        df = analyze_all_months_performance(tickers, LOOKBACK_YEARS)

        if df.empty:
            messagebox.showwarning("No Data", "No valid data found for given tickers.")
            status_label.config(text="Ready", fg="green")
            return

        # -------- FORMAT DATAFRAME --------
        df.set_index("Ticker", inplace=True)
        df = df.round(2)
        df.sort_index(inplace=True)

        ordered_columns = []
        for month in MONTH_NAMES:
            ordered_columns.append(f"{month} Avg. Rtn (%)")
            ordered_columns.append(f"{month} Hit Rate (%)")

        df = df[ordered_columns]

        # -------- SAVE EXCEL --------
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"stock_seasonality_results_{timestamp}.xlsx"
        # Save to user's Documents folder (safe & permission-free)
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        save_path = os.path.join(documents_path, filename)


        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Seasonality", index=True)
            worksheet = writer.sheets["Seasonality"]

            for column_cells in worksheet.columns:
                max_len = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = max_len + 2

        os.startfile(save_path)

        messagebox.showinfo(
            "Completed",
            f"Analysis completed successfully!\n\nExcel file saved & opened:\n{filename}"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))

    status_label.config(text="Ready", fg="green")

# ---------------- GUI SETUP ----------------
root = tk.Tk()
root.title("Stock Seasonality Analyzer")
root.geometry("640x480")
root.resizable(False, False)

tk.Label(
    root,
    text="📊 Stock Seasonality Analyzer",
    font=("Arial", 15, "bold")
).pack(pady=10)

tk.Label(
    root,
    text="Enter stock tickers (comma separated):",
    font=("Arial", 10)
).pack()

stock_entry = tk.Entry(root, width=55, font=("Arial", 11))
stock_entry.pack(pady=5)
stock_entry.insert(0, "AAPL, MSFT, TSLA")

# -------- Popular Tickers Section --------
popular_frame = tk.LabelFrame(
    root,
    text="Popular Tickers",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=5
)
popular_frame.pack(fill="x", padx=15, pady=10)

tk.Label(
    popular_frame,
    text="US Stocks:\n" + ", ".join(POPULAR_TICKERS["US Stocks"]),
    justify="left",
    anchor="w",
    wraplength=580,      # ✅ forces proper wrapping
    font=("Arial", 9)
).pack(anchor="w", fill="x")


tk.Label(
    popular_frame,
    text="\nIndian Stocks (NSE):\n" + ", ".join(POPULAR_TICKERS["Indian Stocks (NSE)"]),
    justify="left",
    anchor="w",
    wraplength=580,
    font=("Arial", 9)
).pack(anchor="w", fill="x")


tk.Button(
    root,
    text="Run Analysis",
    font=("Arial", 12),
    width=22,
    command=run_analysis
).pack(pady=12)

status_label = tk.Label(root, text="Ready", font=("Arial", 10), fg="green")
status_label.pack(pady=8)

root.mainloop()
