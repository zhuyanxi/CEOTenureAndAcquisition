import read_excel as re
import write_excel as we
from operator import itemgetter


# if 'CEO' in app['D0201b'] or '董事长' in app['D0201b'] or '总经理' in app['D0201b'] or '总裁' in app['D0201b']:
#     if '副' not in app['D0201b']:
#         resultlist.append(app)
#         print(app)

def main():
    # tables = []
    # for i in range(1, 10):
    #     table = re.excel_table_byname(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
    #     tables.extend(table)
    # print(len(tables))
    # we.write_excel('CEO_Final_data.xls', 'final_data', tables)

    # tables = re.subsample_filter(file='CEO_Final_data.xls', by_name=u'final_data')
    # print(len(tables))
    # we.write_excel('CEO_subsample.xls', 'sub_data', tables)
    # tables = re.excel_table_byname(file='CEO_Final_data.xls', by_name=u'final_data')
    # print(len(tables))

    # filter_merge_acquisition()

    # filter_company_industry()

    combine_industry_and_acquisition()

    # filter_ceo_data()

    # sort_ceo_final_data()


def filter_ceo_data():
    tables = []
    for i in range(1, 10):
        # table = re.excel_table_byname(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
        table = re.excel_table_to_OrderedDict(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
        print(len(table))
        tables.extend(table)
    print('len of tables: %d' % len(tables))

    result = []
    for item in tables:
        zhiwei = str(item['D0201a'])
        isCM = zhiwei[0:2]
        isGM = [zhiwei[2:4], zhiwei[4:6]]
        if isCM == '11' or '30' in isGM or '32' in isGM or '37' in isGM:
            if '副' not in item['D0201b']:
                print(item)
                result.append(item)

    print(len(result))
    we.write_excel('CEO_final_V2.xls', 'data', result)


def sort_ceo_final_data():
    table = re.excel_table_to_OrderedDict(file='CEO_final_V2.xls', by_name=u'data')
    result = sorted(table, key=itemgetter('Stkcd', 'Reptdt'))
    for item in result:
        print(item)

    we.write_excel('CEO_final_V2.xls', 'data', result)
    print(len(result))


def combine_industry_and_acquisition():
    '''
    匹配合并并购表和行业表
    '''
    company_industry_tables = re.excel_table_to_OrderedDict(file='company_industry_filter_sample.xls', by_name=u'sub_data')
    M_A_subsample_tables = re.excel_table_to_OrderedDict(file='merger_and_acquisition_filter_sample.xls', by_name=u'data')

    print('len of company_industry_tables: %d' % len(company_industry_tables))
    print('len of M_A_subsample_tables: %d' % len(M_A_subsample_tables))

    result = []
    for i in range(len(M_A_subsample_tables)):
        row_M_A = M_A_subsample_tables[i]
        for j in range(len(company_industry_tables)):
            row_C_I = company_industry_tables[j]
            if row_M_A['Symbol'] == row_C_I['Stkcd']:
                row_M_A.update(row_C_I)
                print(row_M_A)
                result.append(row_M_A)
                break;

    we.write_excel('M&A_industry_combine.xls', 'combine_data', result)
    print(len(result))


def filter_company_industry():
    '''
    过滤公司行业数据：
    剔除B股、ST股、金融行业（行业代码已J开头）
    '''
    tables = re.excel_table_to_OrderedDict(file='company-industry.xls', by_name=u'CG_Co')
    print(len(tables))
    result = []
    for i in range(len(tables)):
        row = tables[i]
        if 'B' not in row['Stktype'] and 'ST' not in row['Stknme'] and 'J' not in row['Nnindcd']:
            print(row)
            result.append(row)

    we.write_excel('company_industry_filter_sample.xls', 'sub_data', result)
    print(len(result))


def filter_merge_acquisition():
    '''
    过滤并购数据：
    BusinessID [业务编码]：包含A=资产、B=股权，
    TradingPositionID [上市公司交易地位编码]：S3101=买方、S3106=既是买方又是标的方
    '''
    tables = re.excel_table_to_OrderedDict(file='merger_and_acquisition.xls', by_name=u'STK_MA_TRADINGMAIN')
    tables.extend(re.excel_table_to_OrderedDict(file='merger_and_acquisition.xls', by_name=u'STK_MA_TRADINGMAIN1'))
    print(len(tables))
    result = []
    for i in range(len(tables)):
        row = tables[i]
        if 'A' in row['BusinessID'] or 'B' in row['BusinessID']:
            if 'S3101' in row['TradingPositionID'] or 'S3106' in row['TradingPositionID']:
                print(row)
                result.append(row)

    we.write_excel('merger_and_acquisition_filter_sample.xls', 'data', result)
    print(len(result))


def filter_dict_list_data(data_table, filter_col, filter):
    result = []
    for i in range(len(data_table)):
        row = data_table[i]
        if filter not in row[filter_col]:
            result.append(row)
    return result


if __name__ == "__main__":
    main()
