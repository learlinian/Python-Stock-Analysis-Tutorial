from openpyxl import load_workbook
import time
import get_stock_info   # import the python file we wrote previously


# check the first column numbers which is empty
def check():
    i = 1
    while sheet['C' + str(i)].value is not None:
        i += 1
    return i


# put out stock information if it is not updated in xlsx file
def put_stock():
    global stock
    if sheet['C1'].value is None:
        index = 'C'
        for key in stock:
            sheet[index + '1'].value = key
            index = chr(ord(index)+1)
        wb.save("C:\python\stock monitor\youtube\stock.xlsx")   # change the path to your file's location


# put time and stock information
def put_info():
    global stock
    column_index = check()
    sheet['A'+str(column_index)] = clock
    init_portfolio = 0
    current_portfolio = 0
    index = 'C'

    # put titles for each columns
    sheet['B' + str(column_index)].value = 'current Value'
    sheet['B' + str(column_index + 1)].value = 'Percentage Change'
    sheet['B' + str(column_index + 2)].value = 'Price Change'
    sheet['B' + str(column_index + 3)].value = 'Portfolio Change'

    # Do calculation with updated stock information
    for key, value in stock.items():
        init_portfolio += value[1]*value[2]
        current_portfolio += value[1] * value[0]
        percentage_change = round((value[0] - value[2])/value[2]*100, 2)
        price_change = value[0]-value[2]

        sheet[index + str(column_index)] = value[0]
        sheet[index + str(column_index + 1)] = str(percentage_change) + '%'
        sheet[index + str(column_index + 2)] = price_change
        index = chr(ord(index)+1)

    portfolio_value_change = current_portfolio - init_portfolio
    portfolio_percentage_change = round(portfolio_value_change/init_portfolio*100, 2)
    sheet['C' + str(column_index + 3)] = portfolio_value_change
    sheet['D' + str(column_index + 3)] = str(portfolio_percentage_change) + '%'


if __name__ == "__main__":
    currentDate = time.strftime('%x')
    currentTime = time.strftime('%X')
    clock = currentDate + ' ' + currentTime
    stock = get_stock_info.get_info()   # get updated information from get_stock_info library

    wb = load_workbook("C:\python\stock monitor\youtube\stock.xlsx")
    sheet = wb.worksheets[0]
    put_stock()
    put_info()
    wb.save("C:\python\stock monitor\youtube\stock.xlsx")   # change the path to your file's location
    
