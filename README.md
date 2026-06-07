# 📊 Bottom-Up Beta Calculator

A free, open-source Streamlit web application that automates the extraction of
historical closing prices from **Yahoo Finance**, assembles them into a
consolidated Excel workbook, and flags any data-quality issues — giving you
everything needed for bottom-up beta estimation.

---

## ✨ Features

| Feature | Detail |
|---|---|
| **Multi-company input** | Add / remove as many tickers as you need |
| **19 built-in indices** | S&P 500, NIFTY 50, FTSE 100, DAX, and more — or enter a custom ticker |
| **Flexible period** | 1 – 30 years back from any chosen valuation date |
| **Frequency choice** | Daily · Weekly · Monthly |
| **Consolidated Excel output** | Sheet 1: Date + Index + all companies; Sheet 2: QC / error report |
| **Data-count deviation flags** | Automatically flags any ticker whose data count deviates >10% from the index |
| **Error log** | Records every ticker that failed (not found, delisted, partial data, etc.) |

---

## 🖥️ Live Demo (Streamlit Cloud)

> Deploy your own free instance in 3 clicks — see **Deployment** below.

---

## 📂 Repository Structure

```
beta-calculator/
├── app.py            ← Main Streamlit application
├── requirements.txt  ← Python dependencies
└── README.md         ← This file
```

---

## 🚀 Run Locally

### 1. Clone the repo
```bash
git clone https://github.com/<YOUR-USERNAME>/beta-calculator.git
cd beta-calculator
```

### 2. Create a virtual environment (recommended)
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Launch
```bash
streamlit run app.py
```

The app opens automatically at **http://localhost:8501**.

---

## ☁️ Deploy for Free on Streamlit Community Cloud

Anyone with a GitHub account can host this app publicly at no cost.

1. **Fork / push** this repo to your own GitHub account.
2. Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub.
3. Click **"New app"**.
4. Select your repository, branch (`main`), and set **Main file path** to `app.py`.
5. Click **"Deploy!"** — your public URL is ready in ~60 seconds.

> Share the URL with anyone; no installation required on their end.

---

## 📖 How to Use

### Step 1 — Enter company tickers
- Type each company's **Yahoo Finance ticker symbol** (e.g. `AAPL`, `RELIANCE.NS`, `TCS.NS`).
- Click **➕ Add company** to add more rows.
- Not sure of a ticker? Search at [finance.yahoo.com](https://finance.yahoo.com) and copy the symbol shown at the top of the page.

### Step 2 — Set parameters
| Field | Description |
|---|---|
| **Benchmark Index** | The index to regress companies against |
| **Valuation Date** | End date of the historical window |
| **Number of Years** | How many years of history to pull |
| **Frequency** | Daily (252 pts/yr) · Weekly (52) · Monthly (12) |

### Step 3 — Click "Fetch Data & Generate Excel"
The app downloads data from Yahoo Finance and produces a downloadable Excel file.

---

## 📊 Excel Output Format

### Sheet 1 — Historical Closing Prices
| Column | Content |
|---|---|
| A | Date (YYYY-MM-DD) |
| B | Index closing price (header = index name + ticker) |
| C onwards | Each company's closing price (header = ticker) |

Alternating row shading, freeze panes at row 3, and 4-decimal formatting are applied automatically.

### Sheet 2 — Error & QC Report
- **Section A** — Run parameters (date, period, frequency, index used, etc.)
- **Section B** — Data count table with deviation flags
  - ✅ OK — count within ±10% of reference
  - ⚠️ DEVIATION — count differs by more than 10%
  - ❌ NO DATA — no data retrieved at all
- **Section C** — Full error log (ticker not found, data unavailable, etc.)

---

## 🌐 Supported Exchanges (sample tickers)

| Exchange | Example Tickers |
|---|---|
| NYSE / NASDAQ (USA) | `AAPL`, `MSFT`, `GOOGL`, `JPM` |
| NSE India | `RELIANCE.NS`, `TCS.NS`, `INFY.NS` |
| BSE India | `RELIANCE.BO`, `TCS.BO` |
| London (LSE) | `BP.L`, `HSBA.L`, `SHEL.L` |
| Frankfurt (XETRA) | `SAP.DE`, `BMW.DE`, `SIE.DE` |
| Hong Kong | `0700.HK`, `9988.HK` |
| Tokyo | `7203.T` (Toyota), `6758.T` (Sony) |
| Toronto | `RY.TO`, `TD.TO` |

---

## 🔑 Common Index Tickers

| Index | Ticker |
|---|---|
| S&P 500 | `^GSPC` |
| NASDAQ Composite | `^IXIC` |
| NIFTY 50 | `^NSEI` |
| BSE SENSEX | `^BSESN` |
| FTSE 100 | `^FTSE` |
| DAX | `^GDAXI` |
| Nikkei 225 | `^N225` |
| Hang Seng | `^HSI` |

---

## 🛠️ Tech Stack

- **[Streamlit](https://streamlit.io)** — web UI framework
- **[yfinance](https://github.com/ranaroussi/yfinance)** — Yahoo Finance data
- **[pandas](https://pandas.pydata.org)** — data manipulation
- **[openpyxl](https://openpyxl.readthedocs.io)** — Excel generation
- **[python-dateutil](https://dateutil.readthedocs.io)** — date arithmetic

---

## ⚠️ Disclaimer

This tool is for **educational and analytical purposes only**. Data is sourced
from Yahoo Finance and may be subject to errors or delays. Always verify data
before using it in investment decisions. This is not financial advice.

---

## 📄 License

MIT License — free to use, modify, and distribute.
