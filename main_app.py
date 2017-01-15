import read_excel as re
import write_excel as we
import common_lib
from operator import itemgetter
from collections import Counter
from itertools import groupby
from collections import OrderedDict
import time


def main():
    # table = re.excel_table_to_OrderedDict(file='M&A_industry_combine.xls', by_name=u'combine_data')
    # years=[]
    # for item in table:
    #     years.append(str(item['FirstDeclareDate']).split('-')[0])
    #
    # print(len(years))
    # count=Counter(years)
    # print(count)

    # companyNums = []
    # for i in range(5):
    #     table = re.excel_table_to_OrderedDict_byIndex(file='CG_Ceo.xls', by_index=i, dataRow=3)
    #     resultList=group_data(table,'Stkcd','StockId')
    #     print(len(resultList))
    #     companyNums.append(len(resultList))
    # print(companyNums)

    # table = re.excel_table_to_OrderedDict_byIndex(file='CG_Ceo.xls', by_index=0, dataRow=3)
    # resultList = group_data(table, 'Stkcd', 'StockId')
    # print(len(resultList))

    # table = re.excel_table_to_OrderedDict_byIndex(file='CG_Ceo.xls', by_index=0, dataRow=3)
    # print(len(table))
    # result = [item for item in table if item['Stkcd'] == '000001']
    # for item in result:
    #     print(dict(item))
    # print(len(result))

    # table = re.excel_table_to_OrderedDict_bySheetName(file='CEO_final_V2.xls', by_name=u'data')
    # result = []
    # for item in table:
    #     zhiwei = str(item['D0201a'])
    #     isGM = [zhiwei[2:4], zhiwei[4:6]]
    #     if '30' in isGM:
    #         result.append(item)
    # print(len(result))
    # we.write_excel('CEO_final_General_Manager.xls', 'data', result)

    # table = re.excel_table_to_OrderedDict_bySheetName(file='company_industry_filter_sample.xls', by_name=u'data')
    # # result = [item for item in table if item['Nnindcd'] == '']
    # result = common_lib.filter_dict_list_equal_or_not(data_table=table, filter_key='Nnindcd', filter_value='', isEqual=True)
    # for item in result:
    #     print(item)
    # print(len(result))
    # we.write_excel('company_industry_can_not_research.xls', 'data', result)


    # table = re.excel_table_to_OrderedDict_bySheetName(file='M&A_industry_combine_except2017.xls',
    #                                                   by_name=u'SucceedAndNonRelevance')
    # for item in table:
    #     item['Year'] = str(item['FirstDeclareDate']).split('-')[0]
    # table.sort(key=itemgetter('Stkcd'))
    # result = []
    # for StockId, items1 in groupby(table, key=itemgetter('Stkcd')):
    #     # print(StockId)
    #     for Year, items2 in groupby(items1, key=itemgetter('Year')):
    #         app = OrderedDict()
    #         app['Stkcd'] = StockId
    #         lenItem2 = 0
    #         app['Year'] = Year
    #         # print(Year)
    #         for i in items2:
    #             lenItem2 += 1
    #             # print(' ', i)
    #         app['Times'] = lenItem2
    #         result.append(app)
    # for item in result:
    #     print(item)
    # print(len(result))
    # we.write_excel('M&A_industry_count_by_StkcdAndYear.xls', 'data', result)


    # table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure.xls',
    #                                                   by_name=u'Tenure')
    # result=[item for item in table if not str(item['行业代码']).startswith('J')]
    # common_lib.print_dict_list(result)
    # print(len(result))
    # we.write_excel('GM_Tenure_with_out_J.xls', 'data', result)

    # table = re.excel_table_to_OrderedDict_bySheetName(file='financial_index_data/FI_T1_Filter(12-31).xls',
    #                                                   by_name=u'data')
    # print(len(table))
    # table = sorted(table, key=itemgetter('Stkcd', 'Accper'))
    # result = []
    # for Stkcd, items1 in groupby(table, key=itemgetter('Stkcd')):
    #     for Accper, items2 in groupby(items1, key=itemgetter('Accper')):
    #         app = OrderedDict()
    #         app['Stkcd'] = Stkcd
    #         app['Accper'] = Accper
    #         lenItems2 = 0
    #         tempList = []
    #         for i in items2:
    #             lenItems2 += 1
    #             tempList.append(OrderedDict(i))
    #         if lenItems2 == 1:
    #             app.update(tempList[0])
    #         else:
    #             app.update([j for j in tempList if j['Typrep'] == 'A'][0])
    #         result.append(app)
    # # common_lib.print_dict_list(result)
    # print(len(result))
    # we.write_excel('financial_index_data/FI_T1_Filter(12-31)_Final.xls', 'data', result)

    # table = re.excel_table_to_OrderedDict_bySheetName(file='financial_index_data/FI_T5_Filter(12-31).xls',
    #                                                   by_name=u'data')
    # print(len(table))
    # table = sorted(table, key=itemgetter('Stkcd', 'Accper'))
    # result = []
    # for Stkcd, items1 in groupby(table, key=itemgetter('Stkcd')):
    #     for Accper, items2 in groupby(items1, key=itemgetter('Accper')):
    #         app = OrderedDict()
    #         app['Stkcd'] = Stkcd
    #         app['Accper'] = Accper
    #         tempList = []
    #         for i in items2:
    #             tempList.append(OrderedDict(i))
    #         print(Stkcd)
    #         print(Accper)
    #         print(tempList)
    #         if len(tempList) == 1:
    #             app.update(tempList[0])
    #         else:
    #             app.update([j for j in tempList if j['Typrep'] == 'A'][0])
    #         result.append(app)
    # # common_lib.print_dict_list(result)
    # print(len(result))
    # we.write_excel('financial_index_data/FI_T5_Filter(12-31)_Final.xls', 'data', result)

    # FI_T1_table = re.excel_table_to_OrderedDict_bySheetName(file='financial_index_data/FI_T1_Filter(12-31)_Final.xls',
    #                                                         by_name=u'data')
    # FI_T5_table = re.excel_table_to_OrderedDict_bySheetName(file='financial_index_data/FI_T5_Filter(12-31)_Final.xls',
    #                                                         by_name=u'data')
    # FI_T10_table = re.excel_table_to_OrderedDict_bySheetName(file='financial_index_data/FI_T10_Filter(12-31)_Final.xls',
    #                                                         by_name=u'data')
    # FI_T1 = [{'Stkcd': item['Stkcd'], 'Accper': item['Accper'], 'Typrep': item['Typrep']} for item in FI_T1_table]
    # FI_T5 = [{'Stkcd': item['Stkcd'], 'Accper': item['Accper'], 'Typrep': item['Typrep']} for item in FI_T5_table]
    # print(FI_T1)
    # print(len(FI_T1))
    # print(len(FI_T5))

    # table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_out_J.xls', by_name=u'data')
    # result = []
    # for StockId, items in groupby(table, key=itemgetter('StockId')):
    #     tempList = []
    #     for i in items:
    #         tempList.append(OrderedDict(i))
    #
    #     tempList.sort(key=itemgetter('Year'))
    #     tempList.reverse()
    #
    #     flag = 0
    #     for i in range(len(tempList)):
    #         if tempList[i]['GM_Name'] == '':
    #             continue
    #         GMName=tempList[i]['GM_Name']
    #         Education=tempList[i]['教育背景']
    #         JRLY=tempList[i]['继任来源']
    #         IsAgent=tempList[i]['是否代理']
    #         for j in range(flag, i):
    #             tempList[j]['GM_Name'] = tempList[i]['GM_Name']
    #             tempList[j]['教育背景'] = tempList[i]['教育背景']
    #             tempList[j]['继任来源'] = tempList[i]['继任来源']
    #             tempList[j]['是否代理'] = tempList[i]['是否代理']
    #         flag = i + 1
    #     tempList.sort(key=itemgetter('Year'))
    #     result.extend(tempList)
    # common_lib.print_dict_list(result)
    # print(len(result))
    # we.write_excel('GM_Tenure_with_out_J_FullVersion.xls', 'data', result)

    # table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_out_J_FullVersion.xls', by_name=u'data')
    # for item in table:
    #     StockId = str(int(item['StockId']))
    #     lenId = len(StockId)
    #     item['StockId'] = StockId if lenId == 6 \
    #         else '0' + StockId if lenId == 5 \
    #         else '00' + StockId if lenId == 4 \
    #         else '000' + StockId if lenId == 3 \
    #         else '0000' + StockId if lenId == 2 \
    #         else '00000' + StockId
    #     item['Year']=str(int(item['Year']))
    #     item['ipo_year'] = str(int(item['ipo_year']))
    #     item['start_year'] = str(int(item['start_year']))
    #     item['tenure'] = str(int(item['tenure']))
    # we.write_excel('GM_Tenure_with_out_J_FullVersion.xls', 'data', table)
    # table.reverse()
    # common_lib.print_dict_list(table)
    # print(len(table))

    # table=[]
    # for i in range(5):
    #     table.extend(re.excel_table_to_OrderedDict_byIndex(file='company_assets_ALL.xls',by_index=i,dataRow=3))
    # table=[item for item in table if str(item['Accper']).endswith('12-31')]
    # we.write_excel('company_assets_Filter(12-31).xls', 'data', table)


    # table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_out_J_FullVersion.xls', by_name=u'data')
    # print(len(table))
    # FirmAssets = re.excel_table_to_OrderedDict_bySheetName(file='company_assets_Filter(12-31).xls', by_name=u'data')
    # print(len(FirmAssets))
    # FinancialIndex = re.excel_table_to_OrderedDict_bySheetName(
    #     file='financial_index_data/FI_T1_Filter(12-31)_Final.xls', by_name=u'data')
    # print(len(FinancialIndex))
    # FirmAge = re.excel_table_to_OrderedDict_bySheetName(file='IPO_Cobasic.xls', by_name=u'IPO_Cobasic')
    # print(len(FirmAge))
    # for item in table:
    #     app = OrderedDict()
    #     asset = [x for x in FirmAssets if
    #              x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
    #     print(asset)
    #     # app['TotalAsset'] = asset[0]['A001000000']
    #     app['TotalAsset'] = float(max(asset, key=itemgetter('A001000000'))['A001000000'])
    #
    #     F_Index = [x for x in FinancialIndex if
    #                x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
    #     print(F_Index)
    #     app['Lev'] = F_Index[0]['F011201A']
    #     app['CashRatio'] = F_Index[0]['F010401A']
    #
    #     Firm_Age = [x for x in FirmAge if x['Stkcd'] == item['StockId']]
    #     year = int(str(Firm_Age[0]['Estbdt']).split('-')[0])
    #     month = int(str(Firm_Age[0]['Estbdt']).split('-')[1])
    #     startYear = year if month <= 6 else year + 1
    #     app['FirmAge'] = str(int(item['Year']) - startYear + 1)
    #
    #     item.update(app)
    # print(table[0])
    # print(len(table))
    # we.write_excel('GM_Tenure_with_Param.xls', 'data', table)

    start = time.clock()
    print(start)

    # CEOAge = []
    # for i in range(9):
    #     CEOAge.extend((re.excel_table_to_OrderedDict_byIndex(file='CG_Director_ALL.xls', by_index=i, dataRow=1)))
    # print(len(CEOAge))
    add_ceoAge_to_GMTenure()

    end = time.clock()
    print(end)
    print('Running time:%s Seconds' % ((end - start)))

    pass


def filter_ceo_data():
    tables = []
    for i in range(1, 10):
        # table = re.excel_table_byname(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
        table = re.excel_table_to_OrderedDict_bySheetName(file='CG_Director_ALL.xls', by_name=u'CEO' + str(i))
        print(len(table))
        tables.extend(table)
    print('len of tables: %d' % len(tables))

    result = []
    CMList = ['10', '11', '13']
    GMList = ['30', '31', '32', '37']
    for item in tables:
        zhiwei = str(item['D0201a'])
        isCM = str(zhiwei[0:2])
        isGM1 = str(zhiwei[2:4])
        isGM2 = str(zhiwei[4:6])
        # if isCM == '11' or '30' in isGM or '32' in isGM or '37' in isGM:
        #     if '副' not in item['D0201b']:
        #         print(item)
        #         result.append(item)
        if isCM in CMList or isGM1 in GMList or isGM2 in GMList:
            # print(item)
            result.append(item)

    print(len(result))
    return result
    # we.write_excel('CEO_final_V3.xls', 'data', result)


def sort_ceo_final_data():
    table = re.excel_table_to_OrderedDict_bySheetName(file='CEO_final_V2.xls', by_name=u'data')
    result = sorted(table, key=itemgetter('Stkcd', 'Reptdt'))
    for item in result:
        print(item)

    we.write_excel('CEO_final_V2.xls', 'data', result)
    print(len(result))


def combine_industry_and_acquisition():
    '''
    匹配合并并购表和行业表
    '''
    company_industry_tables = re.excel_table_to_OrderedDict_bySheetName(file='company_industry_filter_sample.xls',
                                                                        by_name=u'sub_data')
    M_A_subsample_tables = re.excel_table_to_OrderedDict_bySheetName(file='merger_and_acquisition_filter_sample.xls',
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
    剔除B股、ST股、金融行业（行业代码J开头）
    '''
    tables = re.excel_table_to_OrderedDict_bySheetName(file='company-industry.xls', by_name=u'CG_Co')
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
    tables = re.excel_table_to_OrderedDict_bySheetName(file='merger_and_acquisition.xls', by_name=u'STK_MA_TRADINGMAIN')
    tables.extend(
        re.excel_table_to_OrderedDict_bySheetName(file='merger_and_acquisition.xls', by_name=u'STK_MA_TRADINGMAIN1'))
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
    table = re.excel_table_to_OrderedDict_bySheetName(file='M&A_industry_combine_except2017.xls', by_name=u'data')
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
    table = re.excel_table_to_OrderedDict_bySheetName(file='M&A_industry_combine_except2017.xls', by_name=u'data')
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
        # print(group_col)
        lenItem = 0
        for i in items:
            lenItem += 1
            # print(' ', i)
        app['Times'] = lenItem
        result.append(app)
        # print(' ', lenItem)
    return result


def filter_FI_12_31():
    table = []
    for i in range(4):
        table.extend(
            re.excel_table_to_OrderedDict_byIndex(file='financial_index_data/FI_T1_ALL.xls', by_index=i, dataRow=3))
    table = [item for item in table if str(item['Accper']).endswith('12-31')]
    # common_lib.print_dict_list(table)
    print(len(table))
    print(len([item for item in table if str(item['Typrep']) == 'A']))
    print(len([item for item in table if str(item['Typrep']) == 'B']))
    we.write_excel('financial_index_data/FI_T1_Filter(12-31).xls', 'data', table)

    table = []
    for i in range(4):
        table.extend(
            re.excel_table_to_OrderedDict_byIndex(file='financial_index_data/FI_T5_ALL.xls', by_index=i, dataRow=3))
    table = [item for item in table if str(item['Accper']).endswith('12-31')]
    # common_lib.print_dict_list(table)
    print(len(table))
    print(len([item for item in table if str(item['Typrep']) == 'A']))
    print(len([item for item in table if str(item['Typrep']) == 'B']))
    we.write_excel('financial_index_data/FI_T5_Filter(12-31).xls', 'data', table)

    table = []
    for i in range(2):
        table.extend(
            re.excel_table_to_OrderedDict_byIndex(file='financial_index_data/FI_T10_ALL.xls', by_index=i, dataRow=3))
    table = [item for item in table if str(item['Accper']).endswith('12-31')]
    # common_lib.print_dict_list(table)
    print(len(table))
    we.write_excel('financial_index_data/FI_T10_Filter(12-31).xls', 'data', table)


def add_ceoAge_to_GMTenure():
    table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param.xls', by_name=u'data')
    CEOAge = []
    CEOAge = filter_ceo_data()
    print(len(CEOAge))
    for item in table:
        app = OrderedDict()
        # Ceo_Age = [x for x in CEOAge if
        #            x['Stkcd'] == item['StockId'] and str(x['Reptdt']).split('-')[0] == item['Year'] and x['D0101b'] ==
        #            item['GM_Name']]
        if item['GM_Name']!='':
            Ceo_Age = list(filter(lambda x: x['Stkcd'] == item['StockId'] \
                                            and str(x['Reptdt']).split('-')[0] == item['Year'] \
                                            and x['D0101b'] == item['GM_Name'], CEOAge))
            print(Ceo_Age)
            app['CEOAge'] = Ceo_Age[0]['D0401b']
            app['CEOSex'] = 1 if Ceo_Age[0]['D0301b'] == '男' else 0
        else:
            app['CEOAge'] = ''
            app['CEOSex'] = ''
        item.update(app)
    print(table[0])
    print(len(table))
    we.write_excel('GM_Tenure_with_Param(WithCeoAge).xls', 'data', table)


def testRunTime():
    a = sum(range(500000))
    print(a)


if __name__ == "__main__":
    main()
