import urllib.request

# get the stock you want to search
url = "https://www.google.com.sg/search?q="
# url = "http://nianliblog.com"

stock = {"GOOGL": [0.0, 10, 900], "AMZN": [0.0, 5, 1600], "MSFT": [0.0, 100, 90]}

# Try to login google page with domains
headers = {}
headers['User-Agent'] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"


# Parse the number in string format into digital format
def parse(data):
    data = data.replace(',', '')
    return int(data.partition('.')[0])+int(data.partition('.')[2])/100


# Happy collection time
def get_info():
    global stock
    for key in stock:
        url_searching = url + key
        # print(url_searching)
        req = urllib.request.Request(url_searching, headers=headers)
        info = urllib.request.urlopen(req)
        string = str(info.read())
        temp = '\\"' + key + '\\'

        string = string.partition(temp)[2]
        string = string.partition('\\\\"')[2]
        string = string.partition('\\\\"')[0]
        price = parse(string)
        stock[key][0] = price
    return stock


if __name__ == "__main__":
    result = get_info()
    print(result)

