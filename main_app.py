import read_excel as re
import write_excel as we
import mongoDbHelper as mdb
import common_lib
from operator import itemgetter
from collections import Counter
from itertools import groupby
from collections import OrderedDict
import time
import RWTXT


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
    # table=[item for item in table if str(item['Accper']).endswith('01-01')]
    # we.write_excel('company_assets_Filter(01-01).xls', 'data', table)

    # table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param(WithCeoAge_VFinal2).xls',
    #                                                   by_name=u'data')
    # print(len(table))
    # # FirmAssets = re.excel_table_to_OrderedDict_bySheetName(file='company_assets_Filter(12-31).xls', by_name=u'data')
    # # print(len(FirmAssets))
    # FinancialIndex = re.excel_table_to_OrderedDict_bySheetName(
    #     file='financial_index_data/FI_T1_Filter(12-31)_Final.xls', by_name=u'data')
    # print(len(FinancialIndex))
    # FinancialIndex2 = re.excel_table_to_OrderedDict_bySheetName(
    #     file='financial_index_data/FI_T5_Filter(12-31)_Final.xls', by_name=u'data')
    # print(len(FinancialIndex2))
    # FinancialIndex3 = re.excel_table_to_OrderedDict_bySheetName(
    #     file='financial_index_data/FI_T10_Filter(12-31)_Final.xls', by_name=u'data')
    # print(len(FinancialIndex3))
    # # FirmAge = re.excel_table_to_OrderedDict_bySheetName(file='IPO_Cobasic.xls', by_name=u'IPO_Cobasic')
    # # print(len(FirmAge))
    # for item in table:
    #     app = OrderedDict()
    #     # asset = [x for x in FirmAssets if
    #     #          x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
    #     # print(asset)
    #     # app['TotalAsset'] = float(max(asset, key=itemgetter('A001000000'))['A001000000'])
    #     F_Index = [x for x in FinancialIndex if
    #                x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
    #     print(F_Index)
    #     F_Index2 = [x for x in FinancialIndex2 if
    #                 x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
    #     print(F_Index2)
    #     F_Index3 = [x for x in FinancialIndex3 if
    #                 x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
    #     print(F_Index3)
    #
    #     if len(F_Index) == 0 or len(F_Index2) == 0 or len(F_Index3) == 0:
    #         continue
    #     app['EM权益乘数'] = F_Index[0]['F011601A']
    #     app['EDRatio权益对负债比率'] = F_Index[0]['F011801A']
    #     app['DMRatio负债与权益市价比率'] = F_Index[0]['F012501B']
    #     app['ROE净资产收益率A'] = F_Index2[0]['F050501B']
    #     app['EBIT息税前利润'] = F_Index2[0]['F050601B']
    #     app['MarketValueA'] = F_Index3[0]['F100801A']
    #     app['PE1'] = F_Index3[0]['F100101B']
    #     app['PE2'] = F_Index3[0]['F100102B']
    #     app['TobinQ_A'] = F_Index3[0]['F100901A']
    #     app['BTM账面市值比A'] = F_Index3[0]['F100101B']
    #     app['普通股获利率A'] = F_Index3[0]['F101201B']
    #     app['企业价值倍数'] = F_Index3[0]['F101301B']
    #     app['企业价值倍数TTM'] = F_Index3[0]['F101302C']
    #     # Firm_Age = [x for x in FirmAge if x['Stkcd'] == item['StockId']]
    #     # year = int(str(Firm_Age[0]['Estbdt']).split('-')[0])
    #     # month = int(str(Firm_Age[0]['Estbdt']).split('-')[1])
    #     # startYear = year if month <= 6 else year + 1
    #     # app['FirmAge'] = str(int(item['Year']) - startYear + 1)
    #     item.update(app)
    # print(table[0])
    # print(len(table))
    # we.write_excel('GM_Tenure_with_Param(WithCeoAge_VFinal4).xls', 'data', table)

    # start = time.clock()
    # print(start)
    # # add_ceoAge_to_GMTenure()
    # # add_ceoAge_to_GMTenure_V2()
    # # add_ceoAge_to_GMTenure_V3()
    # # add_ceoAge_to_GMTenure_V4()
    # # testRunTime()
    # # calc_CEOAge()
    # table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param(WithCeoAge_VFinal).xls', by_name=u'data')
    # table=[x for x in table if x['GM_Name']!='' and x['CEOAge']=='']
    # common_lib.print_dict_list(table)
    # print(len(table))
    # end = time.clock()
    # print(end)
    # print('Running time:%s Seconds' % ((end - start)))


    # add_Assets0101_to_GMTenure()

    # merge_M_A_by_SameTime()

    # calc_M_A_Times()
    # add_M_A_to_GMTenure()
    # add_ceoAge_to_GMTenure_V4()
    deal_change_tenure_times()
    pass


def filter_ceo_data():
    tables = []
    for i in range(1, 10):
        # table = re.excel_table_byname(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
        table = re.excel_table_to_OrderedDict_bySheetName(file='CG_Director_ALL.xls', by_name=u'CEO' + str(i))
        print(len(table))
        tables.extend(table)
    print('len of tables: %d' % len(tables))
    return tables

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


def calc_CEOAge():
    CEOAge = RWTXT.txt_to_dict_list('CG_Director_ALL.txt')
    print(len(CEOAge))

    result = [OrderedDict({'BirthYear': str(int(x['Reptdt'].split('-')[0]) - int(x['D0401b'][:-2])),
                           'StockId': x['Stkcd'], 'Name': x['D0101b'],
                           'Sex': '1' if x['D0301b'] == '男' else '0'}) for x in CEOAge if x['D0401b'] != '']
    # for i in range(len(CEOAge)):
    #     item = CEOAge[i]
    #     if item['D0401b'] == '':
    #         continue
    #
    #     start = time.clock()
    #
    #     isContainItem = OrderedDict()
    #     isContain = False
    #     StockId = str(item['Stkcd'])
    #     Name = str(item['D0101b'])
    #     BirthYear = str(int(item['Reptdt'].split('-')[0]) - int(item['D0401b'][:-2]))
    #     # isContain = [x for x in result if x['StockId'] == StockId and x['Name'] == Name]
    #     isContainItem['StockId'] = StockId
    #     isContainItem['Name'] = Name
    #     if isContainItem in result:
    #         isContain = True
    #
    #     end = time.clock()
    #     print('Running time:%s Seconds' % (end - start), BirthYear, i)
    #
    #     if isContain:
    #         continue
    #
    #     app = OrderedDict()
    #     app['StockId'] = StockId
    #     app['Name'] = Name
    #     app['BirthYear'] = BirthYear
    #     result.append(app)

    result = list(common_lib.dedupe(result, key=lambda x: (x['StockId'], x['Name'])))

    print(result[123])
    print(len(result))
    result.sort(key=itemgetter('StockId', 'Name'))
    RWTXT.write_dict_list_to_txt('CG_Director_Age.txt', result)


def add_ceoAge_to_GMTenure():
    table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param.xls', by_name=u'data')
    CEOAge = RWTXT.txt_to_dict_list('CG_Director_ALL.txt')
    CEOAge.sort(key=itemgetter('Stkcd'), reverse=True)
    lenCEO = len(CEOAge)
    print(len(CEOAge))
    CEOAge1 = CEOAge[0:int(lenCEO / 2)]
    CEOAge2 = CEOAge[int(lenCEO / 2):int(lenCEO)]
    for item in table:
        app = OrderedDict()

        startT = time.clock()
        Ceo_Age = []
        if item['GM_Name'] != '':
            Ceo_Age = common_lib.Group_Search_dictList(CEOAge, item['StockId'], item['Year'], item['GM_Name'])
        if len(Ceo_Age) > 0:
            app['CEOAge'] = Ceo_Age[0]['D0401b']
            app['CEOSex'] = 1 if Ceo_Age[0]['D0301b'] == '男' else 0
        else:
            app['CEOAge'] = ''
            app['CEOSex'] = ''
        print(Ceo_Age)
        endT = time.clock()
        print('Running time:%s Seconds' % ((endT - startT)))
        item.update(app)
    print(table[0])
    print(len(table))
    we.write_excel('GM_Tenure_with_Param(WithCeoAge_V2).xls', 'data', table)


def add_ceoAge_to_GMTenure_V2():
    GM_Table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param.xls', by_name=u'data')
    CEOAge = filter_ceo_data()
    CEOAge.sort(key=itemgetter('Stkcd'))
    lenCEO = len(CEOAge)
    print(lenCEO)
    CEOAge1 = CEOAge[0:int(lenCEO / 2)]
    CEOAge2 = CEOAge[int(lenCEO / 2):int(lenCEO)]

    for item in GM_Table:
        app = OrderedDict()

        startT = time.clock()

        # Ceo_Age = [x for x in CEOAge1 if x['Stkcd'] == item['StockId'] \
        #            and str(x['Reptdt']).split('-')[0] == item['Year'] \
        #            and x['D0101b'] == item['GM_Name']]
        # if len(Ceo_Age) > 0:
        #     app['CEOAge'] = Ceo_Age[0]['D0401b']
        #     app['CEOSex'] = 1 if Ceo_Age[0]['D0301b'] == '男' else 0
        # else:
        #     Ceo_Age = [x for x in CEOAge2 if x['Stkcd'] == item['StockId'] \
        #                and str(x['Reptdt']).split('-')[0] == item['Year'] \
        #                and x['D0101b'] == item['GM_Name']]
        #     if len(Ceo_Age) > 0:
        #         app['CEOAge'] = Ceo_Age[0]['D0401b']
        #         app['CEOSex'] = 1 if Ceo_Age[0]['D0301b'] == '男' else 0
        #     else:
        #         app['CEOAge'] = ''
        #         app['CEOSex'] = ''

        if item['StockId'] == '600002':
            a = 1
            b = 2

        Ceo_Age = []
        if item['GM_Name'] != '':
            Ceo_Age = common_lib.Binary_Search_dictList(CEOAge, item['StockId'], item['Year'], item['GM_Name'])
        if len(Ceo_Age) > 0:
            app['CEOAge'] = Ceo_Age[0]['D0401b']
            app['CEOSex'] = 1 if Ceo_Age[0]['D0301b'] == '男' else 0
        else:
            app['CEOAge'] = ''
            app['CEOSex'] = ''

        print(Ceo_Age)
        endT = time.clock()
        print('Running time:%s Seconds' % ((endT - startT)))

        item.update(app)
    print(GM_Table[0])
    print(len(GM_Table))
    we.write_excel('GM_Tenure_with_Param(WithCeoAge).xls', 'data', GM_Table)


def add_ceoAge_to_GMTenure_V3():
    GM_Table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param.xls', by_name=u'data')
    mongoClient = mdb.connect_to_db()
    CG_Director_List = mdb.Get_CG_Director_ALL(mongoClient)

    for item in GM_Table:
        app = OrderedDict()

        startT = time.clock()
        Ceo_Age = []
        if item['GM_Name'] != '':
            Ceo_Age = mdb.Search_CG_Director_ALL(CG_Director_List, item['StockId'], item['Year'], item['GM_Name'])
            # Ceo_Age=CG_Director_List.find({'Stkcd': item['StockId'], 'Reptdt': item['Year']+'-12-31', 'D0101b': item['GM_Name']})
        if len(Ceo_Age) > 0:
            app['CEOAge'] = Ceo_Age[0]['D0401b']
            app['CEOSex'] = 1 if Ceo_Age[0]['D0301b'] == '男' else 0
        else:
            app['CEOAge'] = ''
            app['CEOSex'] = ''
        print(Ceo_Age)
        endT = time.clock()
        print('Running time:%s Seconds' % ((endT - startT)))

        item.update(app)
    print(GM_Table[5])
    print(len(GM_Table))
    # we.write_excel('GM_Tenure_with_Param(WithCeoAge).xls', 'data', GM_Table)


def add_ceoAge_to_GMTenure_V4():
    table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param.xls', by_name=u'data')
    CEOAge = RWTXT.txt_to_dict_list('CG_Director_Age.txt')
    print(len(CEOAge))
    for item in table:
        app = OrderedDict()

        startT = time.clock()
        Ceo_Age = []
        if item['GM_Name'] != '':
            Ceo_Age = [x for x in CEOAge if x['StockId'] == item['StockId'] and x['Name'] == item['GM_Name']]
        if len(Ceo_Age) > 0:
            app['CEOAge'] = str(int(item['Year']) - int(Ceo_Age[0]['BirthYear']))
            app['CEOSex'] = Ceo_Age[0]['Sex']
        else:
            app['CEOAge'] = ''
            app['CEOSex'] = ''
        print(Ceo_Age)
        endT = time.clock()
        print('Running time:%s Seconds' % ((endT - startT)))
        item.update(app)
    print(table[0])
    print(len(table))
    we.write_excel('GM_Tenure_with_Param(WithCeoAge_VFinal3).xls', 'data', table)


def add_Assets0101_to_GMTenure():
    table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param(WithCeoAge_VFinal).xls',
                                                      by_name=u'data')
    print(len(table))

    FirmAssets = re.excel_table_to_OrderedDict_bySheetName(file='company_assets_Filter(01-01).xls', by_name=u'data')
    print(len(FirmAssets))

    for item in table:
        app = OrderedDict()
        asset = [x for x in FirmAssets if
                 x['Stkcd'] == item['StockId'] and str(x['Accper']).split('-')[0] == item['Year']]
        print(asset)
        # app['TotalAsset'] = asset[0]['A001000000']
        app['TotalAsset(01-01)'] = float(max(asset, key=itemgetter('A001000000'))['A001000000'])

        item.update(app)
    print(table[0])
    print(len(table))
    we.write_excel('GM_Tenure_with_Param(WithCeoAge_VFinal).xls', 'data', table)


def add_M_A_to_GMTenure():
    table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param(WithCeoAge_VFinal).xls',
                                                      by_name=u'data')
    print(len(table))

    MACount = re.excel_table_to_OrderedDict_bySheetName(file='M&A_TimeCount.xls', by_name=u'data')
    print(len(MACount))

    for item in table:
        app = OrderedDict()
        count = [x for x in MACount if x['StockId'] == item['StockId'] and x['Year'] == item['Year']]
        print(count)

        if len(count) > 0:
            app['MATimes'] = count[0]['Times']
            app['MATotalExpenseValue'] = count[0]['TotalExpenseValue']
        else:
            app['MATimes'] = ''
            app['MATotalExpenseValue'] = ''

        item.update(app)
    print(table[0])
    print(len(table))
    we.write_excel('GM_Tenure_with_Param(WithCeoAge_VFinal2).xls', 'data', table)


def merge_M_A_by_SameTime():
    table = re.excel_table_to_OrderedDict_bySheetName(file='M&A_MergeTime.xls', by_name=u'SucceedAndNonRelevance')
    print(len(table))

    result = []
    for Symbol, items1 in groupby(table, key=itemgetter('Symbol')):
        for FirstDeclareDate, items2 in groupby(items1, key=itemgetter('FirstDeclareDate')):
            app = OrderedDict()
            app['StockId'] = Symbol
            app['FirstDeclareDate'] = FirstDeclareDate
            app['Year'] = str(FirstDeclareDate).split('-')[0]
            ExpenseValue = 0
            for i in items2:
                ExpenseValue += float(i['ExpenseValue'])
            app['ExpenseValue'] = str(ExpenseValue)
            print(app)
            result.append(app)
    we.write_excel('M&A_MergeTime.xls', 'data', result)


def calc_M_A_Times():
    table = re.excel_table_to_OrderedDict_bySheetName(file='M&A_MergeTime.xls', by_name=u'data')
    print(len(table))

    result = []
    for StockId, items1 in groupby(table, key=itemgetter('StockId')):
        for Year, items2 in groupby(items1, key=itemgetter('Year')):
            app = OrderedDict()
            app['StockId'] = StockId
            app['Year'] = Year
            times = 0
            TotalExpenseValue = 0
            for i in items2:
                times += 1
                TotalExpenseValue += float(i['ExpenseValue'])
            app['Times'] = str(times)
            app['TotalExpenseValue'] = str(TotalExpenseValue)
            print(app)
            result.append(app)
    we.write_excel('M&A_TimeCount.xls', 'data', result)


def deal_change_tenure_times():
    table = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_VFinal.xls', by_name=u'data')
    print(len(table))

    ceo_change_times = []
    tenure_one_year = []
    tenure_under_two_year = []
    for StockId, items1 in groupby(table, key=itemgetter('StockId')):
        change_time = 0
        for GM_Name, items2 in groupby(items1, key=itemgetter('GM_Name')):
            change_time += 1
            temp_list = []
            for i in items2:
                temp_list.append(OrderedDict(i))
            if GM_Name != '' and len(temp_list) == 1:
                tenure_one_year.extend(temp_list)
            if GM_Name != '' and len(temp_list) <= 2:
                tenure_under_two_year.extend(temp_list)
        app_change_time = OrderedDict()
        app_change_time['StockId'] = StockId
        app_change_time['ChangeTime'] = change_time - 1
        ceo_change_times.append(app_change_time)

    print(len(ceo_change_times))
    print(len(tenure_one_year))
    print(len(tenure_under_two_year))

    tenure_without_one_year = [i for i in table if i not in tenure_one_year]
    tenure_without_under_two_year = [i for i in table if i not in tenure_under_two_year]
    print(len(tenure_without_one_year))
    print(len(tenure_without_under_two_year))

    we.write_excel('CEO_Change_Time.xls', 'data', ceo_change_times)
    we.write_excel('GM_Tenure_OneYear.xls', 'data', tenure_one_year)
    we.write_excel('GM_Tenure_UnderTwoYear.xls', 'data', tenure_under_two_year)
    we.write_excel('GM_Tenure_Without_OneYear.xls', 'data', tenure_without_one_year)
    we.write_excel('GM_Tenure_Without_UnderTwoYear.xls', 'data', tenure_without_under_two_year)


def testRunTime():
    a = sum(range(5000000))
    print(a)


if __name__ == "__main__":
    main()
