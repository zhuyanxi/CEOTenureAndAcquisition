import read_excel as re
import write_excel as we


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


def combine_industry_and_acquisition():
    '''
    匹配合并并购表和行业表
    '''
    company_industry_tables = re.excel_table_byname(file='company_industry_sub.xls', by_name=u'sub_data')
    M_A_subsample_tables = re.excel_table_byname(file='M_A_subsample.xls', by_name=u'sub_data')

    print('len of company_industry_tables: %d' % len(company_industry_tables))
    print('len of M_A_subsample_tables: %d' % len(M_A_subsample_tables))

    result = []
    for i in range(len(M_A_subsample_tables)):
        row_M_A = dict(M_A_subsample_tables[i])
        for j in range(len(company_industry_tables)):
            row_C_I = dict(company_industry_tables[j])
            if row_M_A['Symbol'] == row_C_I['Stkcd']:
                row_M_A.update(row_C_I)
                result.append(row_M_A)
                break;

    we.write_excel('M_A_industry_combine.xls', 'combine_data', result)
    print(len(result))


def filter_company_industry():
    '''
    过滤公司行业数据：
    剔除B股、ST股、金融行业（行业代码已J开头）
    '''
    tables = re.excel_table_byname(file='company-industry.xls', by_name=u'CG_Co')
    print(len(tables))
    result = []
    for i in range(len(tables)):
        app = tables[i]
        if 'B' not in app['Stktype'] and 'ST' not in app['Stknme'] and 'J' not in app['Nnindcd']:
            result.append(app)

    we.write_excel('company_industry_sub.xls', 'sub_data', result)
    print(len(result))


def filter_merge_acquisition():
    '''
    过滤并购数据：
    BusinessID [业务编码]：包含A=资产、B=股权，
    TradingPositionID [上市公司交易地位编码]：S3101=买方、S3106=既是买方又是标的方
    '''
    tables = re.excel_table_byname(file='merger_and_acquisition.xls', by_name=u'STK_MA_TRADINGMAIN')
    tables.extend(re.excel_table_byname(file='merger_and_acquisition.xls', by_name=u'STK_MA_TRADINGMAIN1'))
    print(len(tables))
    result = []
    for i in range(len(tables)):
        app = tables[i]
        if 'A' in app['BusinessID'] or 'B' in app['BusinessID']:
            if 'S3101' in app['TradingPositionID'] or 'S3106' in app['TradingPositionID']:
                result.append(app)

    we.write_excel('M_A_subsample.xls', 'sub_data', result)
    print(len(result))


def filter_dict_list_data(data_table, filter_col, filter):
    result = []
    for i in range(data_table):
        row = data_table[i]
        if filter in row[filter_col]:
            result.append(row)
    return result


if __name__ == "__main__":
    main()
