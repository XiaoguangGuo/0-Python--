import requests

def fetch_exchange_rates(app_id, currencies):
    api_url = 'https://openexchangerates.org/api/latest.json?app_id={}&symbols={}'
    rates = {}

    try:
        response = requests.get(api_url.format(app_id, ','.join(currencies)), headers={'Authorization': f'Token {app_id}'})
        response.raise_for_status()
        data = response.json()
        for currency in currencies:
            rate = data['rates'][currency]
            rates[currency] = rate
        print('连接API成功！以下是货币对美元的汇率：')
        print(rates)
    except requests.exceptions.HTTPError as error:
        print(f'连接API出错：{error}')
        
    return rates

if __name__ == '__main__':
    app_id = '438a43ad7170441aa0c7a00caebf086f'
    currencies = ['USD', 'CAD', 'EUR', 'GBP', 'JPY', 'MXN', 'SEK']
    exchange_rates = fetch_exchange_rates(app_id, currencies)
