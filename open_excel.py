import xlrd


def open_excel(file='CG_Director_sample.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))
