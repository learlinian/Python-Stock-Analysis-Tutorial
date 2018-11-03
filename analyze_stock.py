from openpyxl import load_workbook
import get_stock_info  # import the python file we wrote previously
import time


# check the first column numbers which is empty
def check():
    global sheet
    i = 1
    while sheet['C' + str(i)].value is not None:
        i += 1
    return i


# put out stock information if it is not updated in xlsx file
def put_stock():
    global stock
    global sheet
    if sheet['C1'].value is None:
        index = 'C'
        for key in stock:
            sheet[index+'1'].value = key
            index = chr(ord(index)+1)
        wb.save('stock.xlsx')   # change the path to your file's location


# put time and stock information
def put_info():
    global sheet
    global stock
    init_port = 0
    current_port = 0
    row_index = check()

    currentData = time.strftime('%x')
    currentTime = time.strftime('%X')
    clock = currentData + ' ' + currentTime
    sheet['A' + str(row_index)].value = clock

    # put stock features
    sheet['B' + str(row_index)].value = 'current value'
    sheet['B' + str(row_index+1)].value = 'percentage change'
    sheet['B' + str(row_index+2)].value = 'price change'
    sheet['B' + str(row_index+3)].value = 'portfolio change'

    # put stock info
    for key, value in stock.items():
        index = 'C'
        init_price = value[0]
        number = value[2]
        current_price = value[1]
        init_port += init_price*number
        current_port += current_price*number
        price_change = current_price - init_price
        percentage_change = (current_price - init_price)/init_price * 100

        while sheet[index + '1'].value != key:
            index = chr(ord(index)+1)
        sheet[index + str(row_index)].value = current_price
        sheet[index + str(row_index+1)].value = str(round(percentage_change, 2)) + '%'
        sheet[index + str(row_index+2)].value = price_change

    port_price_change = current_port - init_port
    port_percentage_change = (current_port - init_port)/init_port * 100
    sheet['C' + str(row_index+3)].value = port_price_change
    sheet['D' + str(row_index+3)].value = str(round(port_percentage_change, 2)) + '%'


if __name__ == "__main__":
    stock = get_stock_info.get_info()
    wb = load_workbook('stock.xlsx')
    sheet = wb.worksheets[0]
    put_stock()
    put_info()

    wb.save('stock.xlsx')  # save the excel file    
