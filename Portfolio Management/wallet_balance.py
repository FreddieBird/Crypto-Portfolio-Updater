## Accesses wallet balances for Binance, FTX and Etherscan.io

## Balance checker for Binance exchange

import ccxt
import bs4
import re
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as soup

############################## ehterscan address #######################################
etherscan_address = ""
############################## ehterscan address #######################################

################################ FTX API Keys ##########################################
ftx_api_key = ""
ftx_secret_key = ""
################################ FTX API Keys ##########################################

############################## Binance API Keys ########################################
binance_api_key = ""
binance_secret_key = ""
############################## Binance API Keys ########################################


# Connects to the exchanges you have set to be true in the param list
def retrieve_balances(crypto_portfolio, binance=True, ftx=True, etherscan=True):

    # initialise balances to 0
    balances = {}
    for crypto in crypto_portfolio:
        balances[crypto] = 0.0

    if binance:
        binance_ex = ccxt.binance({
            'apiKey': binance_api_key,
            'secret': binance_secret_key,
            'enableRateLimit': True,
        })

        # fetch balances
        binance_balance = binance_ex.fetch_balance()

        # insert balances from binance
        print("Retrieving balances from Binance...")
        for crypto in crypto_portfolio:
            try:
                balances[crypto] += binance_balance['total'][crypto]

            except:
                continue
        print("Successfully retrieved balances from Binance!")

    if ftx:
        ftx_ex = ccxt.ftx({
            'apiKey': ftx_api_key,
            'secret': ftx_secret_key,
            'enableRateLimit': True,
        })

        # fetch balances
        ftx_balance = ftx_ex.fetch_balance()

        # insert balances from ftx
        print("Retrieving balances from FTX...")
        for crypto in crypto_portfolio:
            try:
                balances[crypto] += ftx_balance['total'][crypto]

            except:
                continue
        print("Successfully retrieved balances from FTX!")

    if etherscan:
        url = 'https://etherscan.io/address/' + etherscan_address
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        webpage = urlopen(req).read()
        page_soup = soup(webpage, "html.parser")

        print("Retrieving balances from etherscan...")
        for crypto in crypto_portfolio:
            try:
                # etherscan scrape
                if crypto != "ETH":
                    data = page_soup.find_all('span', string=re.compile(f'\d {crypto}'))
                    bal = (data[0].contents)[0]
                    bal = float(re.findall("\d+\.\d+", bal)[0])
                    balances[crypto] += bal
                else:
                    data = [div.get_text() for div in page_soup.find_all('div', attrs={'class':'col-md-8'})][0]
                    bal = float(re.findall("\d+\.\d+", data)[0])
                    balances["ETH"] += bal
            except:
                continue
        print("Successfully retrieved balances from etherscan!")


    # Overall Balance Summary
    for k in balances:
        print(f"{k}: {balances[k]}")

    return balances

#balances = retrieve_balances(crypto_portfolio)
