import xlwt

def write_excel(file,sheetname,data):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheetname)


    colname = list(data[0].keys())

    for i in range(len(colname)):
        ws.write(0, i, colname[i])

    for k in range(len(data)):
        for j in range(len(data[k])):
            ws.write(k+1, j, data[k][colname[j]])

    wb.save(file)
#if ('CEO' or '总裁' or '董事长' or '总经理') in app['D0201b']: