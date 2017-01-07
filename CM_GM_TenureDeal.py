import read_excel as re
import write_excel as we


def main():
    # deal_CM_GM_Tenure()
    calculate_tenure()


def deal_CM_GM_Tenure():
    CM_GM_table = re.excel_table_byname(file='CM_GM.xls', by_name=u'data')
    print('len of CM_GM_table: %d' % len(CM_GM_table))
    company_industry_tables = re.excel_table_byname(file='company_industry_sub.xls', by_name=u'sub_data')
    print('len of company_industry_tables: %d' % len(company_industry_tables))

    filter_table = []
    for i in range(len(CM_GM_table)):
        row_CM_GM = CM_GM_table[i]
        is_ST = str(row_CM_GM['stockname']).find('ST')
        is_SH_B = str(row_CM_GM['stocknum']).find('900')
        is_SZ_B = str(row_CM_GM['stocknum']).find('200')
        if str(row_CM_GM['stocknum']).find('200') != 0 and str(row_CM_GM['stocknum']).find('900') != 0 and str(
                row_CM_GM['stockname']).find('ST') == -1:
            for j in range(len(company_industry_tables)):
                row_C_I = dict(company_industry_tables[j])
                if row_CM_GM['stocknum'] == row_C_I['Stkcd']:
                    # row_CM_GM.update(row_C_I)
                    filter_table.append(row_CM_GM)
                    break;

    print('len of filter_CM_GM_table: %d' % len(filter_table))

    # output_colnames = ['Year', 'StockNum', 'StockName', 'Chairman', 'CM_Tenure', 'General_Manager', 'GM_Tenure']
    result = []
    for i in range(len(filter_table)):
        row = dict(filter_table[i])
        row_key = list(row.keys())
        app = {'Year': '', 'StockNum': '', 'StockName': '', 'Chairman': '', 'CM_Tenure': '', 'General_Manager': '',
               'GM_Tenure': ''}
        for j in range(len(row_key)):
            if 'Y0801b' in row_key[j]:
                app['StockNum'] = row['stocknum']
                app['StockName'] = row['stockname']

                time = str(row_key[j]).split('=')[1]
                year = str(time).split('-')[0]
                app['Year'] = year
                app['Chairman'] = row[row_key[j]]
                result.append(app)

            if 'Y0901b' in row_key[j]:
                app['StockNum'] = row['stocknum']
                app['StockName'] = row['stockname']

                time = str(row_key[j]).split('=')[1]
                year = str(time).split('-')[0]
                app['Year'] = year
                app['General_Manager'] = row[row_key[j]]
                result.append(app)

            app = {'Year': '', 'StockNum': '', 'StockName': '', 'Chairman': '', 'CM_Tenure': '', 'General_Manager': '',
                   'GM_Tenure': ''}

    print('len of first result: %d' % len(result))

    final_result = []
    isContinue = 0
    for i in result:
        if i['Chairman'] == '' and i['General_Manager'] == '':
            continue
        for j in final_result:
            if i['Year'] == j['Year'] and i['StockNum'] == j['StockNum']:
                j['Chairman'] = i['Chairman'] if i['Chairman'] != '' else j['Chairman']
                j['General_Manager'] = i['General_Manager'] if i['General_Manager'] != '' else j['General_Manager']
                isContinue = 1
                print(j)
                break
        if isContinue == 1:
            isContinue = 0
            continue
        final_result.append(i)

    # for item in final_result:
    #     print(item)

    print('len of final result: %d' % len(final_result))
    we.write_excel('CM_GM_Tenure.xls', 'data', final_result)
    # for k in range(len(result)):
    #     print(result[k])


def calculate_tenure():
    CM_GM_table = re.excel_table_byname(file='CM_GM_Tenure.xls', by_name=u'data')
    print('len of CM_GM_table: %d' % len(CM_GM_table))


if __name__ == "__main__":
    main()
