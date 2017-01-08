import read_excel as re
import write_excel as we
from operator import itemgetter
from collections import Counter
from itertools import groupby
from collections import OrderedDict


def main():
    # filter_merge_acquisition()

    # filter_company_industry()

    # combine_industry_and_acquisition()

    # filter_ceo_data()

    # sort_ceo_final_data()

    # table = re.excel_table_to_OrderedDict(file='M&A_industry_combine.xls', by_name=u'combine_data')
    # years=[]
    # for item in table:
    #     years.append(str(item['FirstDeclareDate']).split('-')[0])
    #
    # print(len(years))
    # count=Counter(years)
    # print(count)
    table = re.excel_table_to_OrderedDict(file='M&A_industry_combine.xls', by_name=u'combine_data')
    result = group_data(table, 'IsSucceed', 'IsSuccess')
    for item in result:
        print(item)


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
    company_industry_tables = re.excel_table_to_OrderedDict(file='company_industry_filter_sample.xls',
                                                            by_name=u'sub_data')
    M_A_subsample_tables = re.excel_table_to_OrderedDict(file='merger_and_acquisition_filter_sample.xls',
                                                         by_name=u'data')

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


def group_M_A_by_year():
    table = re.excel_table_to_OrderedDict(file='M&A_industry_combine.xls', by_name=u'combine_data')
    for item in table:
        item['FirstDeclareYear'] = str(item['FirstDeclareDate']).split('-')[0]

    result = []
    # Sort by the desired field first
    table.sort(key=itemgetter('FirstDeclareYear'))
    # Iterate in groups
    for FirstDeclareYear, items in groupby(table, key=itemgetter('FirstDeclareYear')):
        app = {}
        app['Year'] = FirstDeclareYear
        print(FirstDeclareYear)
        lenItem = 0
        for i in items:
            lenItem += 1
            # print(' ', i)
        app['Times'] = lenItem
        result.append(app)
        print(' ', lenItem)

    we.write_excel('M&A_groupby_year.xls', 'data', result)


def group_M_A_by_industry():
    table = re.excel_table_to_OrderedDict(file='M&A_industry_combine.xls', by_name=u'combine_data')
    result1 = []
    # Sort by the desired field first
    table.sort(key=itemgetter('Nnindcd'))
    # Iterate in groups
    for Nnindcd, items in groupby(table, key=itemgetter('Nnindcd')):
        app = OrderedDict()
        app['IndustryId'] = Nnindcd
        print(Nnindcd)
        lenItem = 0
        industry_name = ''
        for i in items:
            lenItem += 1
            industry_name = dict(i)['Nnindnme']
            # print(' ', i)
        app['Times'] = lenItem
        app['IndustryName'] = industry_name
        result1.append(app)
        print(' ', lenItem, '---' + industry_name)

    we.write_excel('M&A_groupby_industry.xls', 'data', result1)


def group_data(data_table, group_col, group_title):
    result = []
    data_table.sort(key=itemgetter(group_col))
    for group_col, items in groupby(data_table, key=itemgetter(group_col)):
        app = OrderedDict()
        app[group_title] = group_col
        print(group_col)
        lenItem = 0
        for i in items:
            lenItem += 1
            # print(' ', i)
        app['Times'] = lenItem
        result.append(app)
        print(' ', lenItem)
    return result


if __name__ == "__main__":
    main()
