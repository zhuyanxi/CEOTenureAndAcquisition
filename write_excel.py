import xlwt


def write_excel(file, sheetname, data):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheetname)

    colname = list(data[0].keys())

    for i in range(len(colname)):
        ws.write(0, i, colname[i])

    for k in range(len(data)):
        for j in range(len(data[k])):
            ws.write(k + 1, j, data[k][colname[j]])

    wb.save(file)


def write_excel_two_Sheet(file, sheetname, data):
    wb = xlwt.Workbook()
    colname = list(data[0].keys())

    ws1 = wb.add_sheet(sheetname + str(1))
    for j in range(len(colname)):
        ws1.write(0, j, colname[j])
    for k in range(len(data[:65535])):
        for l in range(len(data[k])):
            ws1.write(k + 1, l, data[k][colname[l]])

    ws2 = wb.add_sheet(sheetname + str(2))
    for j in range(len(colname)):
        ws2.write(0, j, colname[j])
    for k in range(65535, len(data[65535:])):
        for l in range(len(data[k])):
            ws2.write(k + 1, l, data[k][colname[l]])

    wb.save(file)
