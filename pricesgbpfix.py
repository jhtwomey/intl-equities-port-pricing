import yfinance as yf
import pandas as pd
from datetime import datetime
import requests

def get_exchange_rates():
    """
    Fetch the latest exchange rates from a reliable API.
    
    Returns:
        dict: Dictionary of exchange rates relative to USD
    """
    try:
        # Using exchangerate-api.com for currency conversion
        response = requests.get('https://v6.exchangerate-api.com/v6/EXCHANGERATE API KEY HERE/latest/USD')
        data = response.json()
        
        if data['result'] == 'success':
            return data['conversion_rates']
        else:
            print("Failed to fetch exchange rates")
            return {}
    except Exception as e:
        print(f"Error fetching exchange rates: {e}")
        return {}

def convert_to_usd(price, ticker, exchange_rates):
    """
    Convert stock price to USD based on the ticker's exchange.
    
    Args:
        price (float): Stock price in local currency
        ticker (str): Stock ticker symbol
        exchange_rates (dict): Dictionary of exchange rates
    
    Returns:
        tuple: (USD price, exchange rate, currency)
    """
    # Dictionary to map ticker suffixes to currencies
    currency_map = {
        '.T': 'JPY',      # Tokyo Stock Exchange (Japanese Yen)
        '.DE': 'EUR',     # German Stock Exchange (Euro)
        '.L': 'GBP',      # London Stock Exchange (British Pound)
        '.MI': 'EUR',     # Milan Stock Exchange (Euro)
        '.CO': 'DKK',     # Copenhagen Stock Exchange (Danish Krone)
        '.AS': 'EUR',     # Amsterdam Stock Exchange (Euro)
        '.ST': 'SEK',     # Stockholm Stock Exchange (Swedish Krona)
        '.PA': 'EUR',     # Paris Stock Exchange (Euro)
        '.BR': 'EUR',     # Brussels Stock Exchange (Euro)
    }
    
    # Default to local currency if no specific mapping
    default_currency = 'USD'
    
    # Find the currency based on ticker suffix
    currency = default_currency
    for suffix, curr in currency_map.items():
        if ticker.endswith(suffix):
            currency = curr
            break
    
    # Special handling for GBP (divide by 100 if from London Stock Exchange)
    if currency == 'GBP':
        price = price / 100
    
    # If currency is already USD, return the price
    if currency == 'USD':
        return price, 1.0, currency
    
    # Convert to USD using exchange rates
    try:
        if currency in exchange_rates:
            exchange_rate = exchange_rates[currency]
            usd_price = price / exchange_rate
            return usd_price, exchange_rate, currency
        else:
            print(f"No exchange rate found for {currency}")
            return None, None, currency
    except Exception as e:
        print(f"Error converting {ticker} from {currency} to USD: {e}")
        return None, None, currency

def get_stock_prices(tickers):
    """
    Fetch the latest closing prices for a list of stock tickers and convert to USD.
    
    Args:
        tickers (list): List of stock ticker symbols
    
    Returns:
        pandas.DataFrame: DataFrame with stock tickers and their latest closing prices in USD
    """
    # Fetch exchange rates
    exchange_rates = get_exchange_rates()
    
    # Create an empty list to store stock data
    stock_data = []
    
    # Fetch data for each ticker
    for ticker in tickers:
        try:
            # Fetch stock information
            stock = yf.Ticker(ticker)
            
            # Get the most recent closing price
            hist = stock.history(period="1d")
            
            if not hist.empty:
                close_price = hist['Close'].iloc[-1]
                
                # Convert to USD and get exchange rate
                usd_price, exchange_rate, currency = convert_to_usd(close_price, ticker, exchange_rates)
                
                stock_data.append({
                    'Ticker': ticker,
                    'Closing Price (Original Currency)': close_price,
                    'Original Currency': currency,
                    'Exchange Rate (to USD)': exchange_rate,
                    'Closing Price (USD)': usd_price,
                    'Date': hist.index[-1].date()
                })
            else:
                print(f"No data found for {ticker}")
                stock_data.append({
                    'Ticker': ticker,
                    'Closing Price (Original Currency)': None,
                    'Original Currency': None,
                    'Exchange Rate (to USD)': None,
                    'Closing Price (USD)': None,
                    'Date': None
                })
        
        except Exception as e:
            print(f"Error fetching data for {ticker}: {e}")
            stock_data.append({
                'Ticker': ticker,
                'Closing Price (Original Currency)': None,
                'Original Currency': None,
                'Exchange Rate (to USD)': None,
                'Closing Price (USD)': None,
                'Date': None
            })
    
    # Convert to DataFrame
    df = pd.DataFrame(stock_data)
    
    return df

def main():
    # List of stock tickers
    tickers = [
        '2802.T', '3064.T', '3697.T', '3994.T', '4478.T', '4755.T', '6027.T', 
        '6501.T', '6544.T', '6857.T', '6920.T', '6951.T', '7011.T', '7747.T', 
        '7979.T', '8306.T', '8316.T', 'ACSO.L', 'ADN1.DE', 'ADYEN.AS', 'AIXA.DE', 
        'ALK-B.CO', 'ASM.AS', 'ASR', 'AWE.L', 'BA.L', 'BC.MI', 'BOKU.L', 
        'BOOZT.ST', 'CRSP', 'DLG.MI', 'FUTR.L', 'ING', 'MBLY', 'MELI', 'MRX.DE', 
        'MTLS', 'NEM.DE', 'OCDO.L', 'ONWD.BR', 'OPRA', 'RAY-B.ST', 'RHM.DE', 
        'SFTBY', 'SHEL', 'SIE.DE', 'SIX2.DE', 'SPOT', 'SU.PA', 'TOBII.ST', 'TSM', 
        'VIT-B.ST', 'WISE.L', 'WOSG.L', 'ZAL.DE'
    ]
    
    # Fetch stock prices
    stock_prices = get_stock_prices(tickers)
    
    # Generate filename with current date and time
    current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"stock_prices_{current_datetime}.xlsx"
    
    # Save to Excel
    stock_prices.to_excel(filename, index=False)
    print(f"Stock prices saved to {filename}")
    
    # Display the DataFrame
    print(stock_prices)

if __name__ == "__main__":
    main()
