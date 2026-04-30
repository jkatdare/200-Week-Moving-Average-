import pandas as pd

def get_tickers() -> list[str]:

    sp500     = pd.read_csv("sp500.csv")["ticker"].tolist()
    nasdaq100 = pd.read_csv("nasdaq100.csv")["ticker"].tolist()
    dji30     = pd.read_csv("dji30.csv")["ticker"].tolist()

    all_tickers    = sp500 + nasdaq100 + dji30
    unique_tickers = list(dict.fromkeys(all_tickers))

    print(f"Loaded {len(sp500)} S&P 500, {len(nasdaq100)} NASDAQ 100, {len(dji30)} DJIA 30 tickers.")
    print(f"{len(unique_tickers)} unique tickers total.")

    return unique_tickers


if __name__ == "__main__":
    tickers = get_tickers()
    print(tickers)
