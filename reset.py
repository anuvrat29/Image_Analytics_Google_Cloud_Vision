import xlwings as xw

def run():
    wb = xw.Book.caller()
    wb.sheets[0].range('E11').value = ""
    wb.sheets[0].range('E13').value = ""
    wb.sheets[0].range('B17').value = ""
    wb.sheets[0].range('B22').value = ""
    wb.sheets[0].range('I11').value = ""
    wb.sheets[0].range('G20').value = ""
    wb.sheets[0].range('L14').value = ""
    wb.sheets[0].range('E27').value = ""
    wb.sheets[0].range('F27').value = ""
    wb.sheets[0].range('G27').value = ""
    wb.sheets[0].range('H27').value = ""
    wb.sheets[0].range('I27').value = ""
    wb.sheets[0].range('J27').value = ""
    wb.sheets[0].range('K27').value = ""
    wb.sheets[0].range('L27').value = ""
    wb.sheets[0].range('M27').value = ""
