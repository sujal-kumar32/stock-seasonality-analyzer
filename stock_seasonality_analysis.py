import yfinance as yf
import pandas as pd
from tqdm import tqdm
from calendar import month_abbr

# --- Configuration ---
TICKERS = [
    'AAPL',
    'MSFT',
    'RTX',
    'UL',
    'JPM',
    'XOM',
    'V',
    'WMT',
]

LOOKBACK_YEARS = 15
MONTH_NAMES = list(month_abbr)[1:]  # Jan to Dec


def analyze_all_months_performance(tickers: list, years: int) -> pd.DataFrame:
    print(f"Starting 12-month seasonality analysis for {len(tickers)} tickers over the last {years} years...")

    end_date = pd.to_datetime('today').normalize()
    start_date = end_date - pd.DateOffset(years=years)

    all_results = []

    for ticker in tqdm(tickers, desc="Processing Tickers"):
        try:
            data = yf.download(ticker, start=start_date, end=end_date, progress=False)

            if data.empty:
                print(f"Warning: No data for {ticker}")
                continue

            ticker_results = {'Ticker': ticker}

            for month_num, month_name in enumerate(MONTH_NAMES, start=1):
                monthly_returns = []

                for year in range(start_date.year, end_date.year + 1):
                    month_data = data[
                        (data.index.month == month_num) &
                        (data.index.year == year)
                    ]

                    if month_data.empty or len(month_data) < 5:
                        continue

                    # ✅ FIX: force scalar values
                    start_price = float(month_data['Open'].iloc[0])
                    end_price = float(month_data['Close'].iloc[-1])

                    monthly_return = (end_price - start_price) / start_price
                    monthly_returns.append(monthly_return)

                if len(monthly_returns) == 0:
                    avg_return = 0.0
                    hit_rate = 0.0
                else:
                    monthly_series = pd.Series(monthly_returns)

                    # ✅ FIX: explicit float conversion
                    avg_return = float(monthly_series.mean()) * 100
                    hit_rate = (monthly_series.gt(0).sum() / len(monthly_series)) * 100

                ticker_results[f'{month_name} Avg. Rtn (%)'] = avg_return
                ticker_results[f'{month_name} Hit Rate (%)'] = hit_rate

            all_results.append(ticker_results)

        except Exception as e:
            print(f"Error processing {ticker}: {e}")
            continue

    results_df = pd.DataFrame(all_results)

    if results_df.empty:
        print("No results to display.")
        return pd.DataFrame()

    results_df.set_index('Ticker', inplace=True)
    return results_df


# ---- Run ----
if __name__ == "__main__":
    df = analyze_all_months_performance(TICKERS, LOOKBACK_YEARS)
    pd.set_option('display.max_columns', None)
    print(df.round(2))


if __name__ == "__main__":
    df = analyze_all_months_performance(TICKERS, LOOKBACK_YEARS)
    pd.set_option('display.max_columns', None)
    print(df.round(2))
    input("\nPress Enter to exit...")

