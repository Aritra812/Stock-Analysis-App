# Stock Analysis App 📈

This is a simple Python-based **Stock Analysis and Prediction** app with a GUI built using **Tkinter**.  
It allows you to:
- Fetch **daily stock prices** from Yahoo Finance  
- View **historical high/low prices** for custom durations  
- Generate **candlestick charts** using `mplfinance`  
- **Predict future stock prices** using a basic Linear Regression model  

---

## Features
- **Fetch Daily Prices** → Get today’s stock Open, High, Low, Close, and Volume  
- **Historical Reports** → See high and low prices for 1 month, 6 months, 1 year, or 5 years  
- **Candlestick Charts** → Visualize stock movement trends  
- **AI Prediction** → Predicts future Open, High, Low, Close, and Volume using past 6 months of data  

---

## Technologies Used
- **Python**
- **Tkinter** (GUI)
- **yfinance** (Yahoo Finance API)
- **mplfinance** (Candlestick plotting)
- **scikit-learn** (Linear Regression model)
- **pandas** & **numpy**
- **openpyxl** (Excel file handling)

---

## How to Run
1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/Stock-Analysis-App.git
   ```
2. Install required libraries:
   ```bash
   pip install yfinance mplfinance pandas numpy scikit-learn openpyxl
   ```
3. Run the app:
   ```bash
   python "stock future prediction.py"
   ```

---

## Disclaimer
This is for **educational purposes only**.  
The prediction model is very simple and should **not** be used for actual trading decisions.  
This is just a **normal AI project**, and I am currently working on a **similar but more efficient AI** that will provide better accuracy and features.
