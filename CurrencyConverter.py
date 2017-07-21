#!/usr/bin/python

# Fixer.io is a free JSON API for current and historical foreign exchange rates.
# It relies on daily feeds published by the European Central Bank.

# Code taken from https://github.com/akora/fixerio-python

import sys
import requests

base_url = 'http://api.fixer.io/latest'
base_currencies = ['EUR', 'GBP', 'CNY', 'MXN', 'AUD']
rate_in = 'USD'


def get_currency_rate(currency, rate_in):
    query = base_url + '?base=%s&symbols=%s' % (currency, rate_in)
    try:
        response = requests.get(query)
        # print("[%s] %s" % (response.status_code, response.url))
        if response.status_code != 200:
            response = 'N/A'
            return response
        else:
            rates = response.json()
            rate_in_currency = rates["rates"][rate_in]
            return rate_in_currency
    except requests.ConnectionError as error:
        logging.error(": ", error)
        print error
        sys.exit(1)

def get_rates():
    currency_list = []
    currency_list.append(float(1))
    for currency in base_currencies:
        rate = get_currency_rate(currency, rate_in)
        currency_list.append(rate)
        logging.info(": 1 " + currency + " = " + str(rate) + " " + rate_in)
        print("1 " + currency + " = " + str(rate) + " " + rate_in)
    return currency_list    

def convert_currency(base_currency, rate_in):
    rate = get_currency_rate(base_currency, rate_in)
    logging.info(": 1 " + base_currency + " = " + str(rate) + " " + rate_in)
    print("1 " + base_currency + " = " + str(rate) + " " + rate_in)
    return rate    

def main():
    logging.info(get_rates())
    print(get_rates())

if __name__ == '__main__':
    main()
    
