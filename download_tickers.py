import requests
from bs4 import BeautifulSoup
import pandas as pd

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

SOURCES = {
    "sp500":     "https://slickcharts.com/sp500",
    "nasdaq100": "https://slickcharts.com/nasdaq100",
    "dji30":     "https://slickcharts.com/dowjones",
}


def scrape_tickers(url):
    html     = requests.get(url, headers=HEADERS).text
    soup     = BeautifulSoup(html, "html.parser")
    table    = soup.find("table")
    tickers  = []
    for row in table.find_all("tr")[1:]:
        cols = row.find_all("td")
        if len(cols) >= 3:
            ticker = cols[2].text.strip().replace(".", "-")
            if ticker:
                tickers.append(ticker)
    return tickers


for name, url in SOURCES.items():
    print(f"Downloading {name} tickers from slickcharts.com...")
    tickers = scrape_tickers(url)
    pd.DataFrame(tickers, columns=["ticker"]).to_csv(f"{name}.csv", index=False)
    print(f"  Saved {len(tickers)} tickers to {name}.csv")

print("\nDone! Run main.py to fetch price data.")
