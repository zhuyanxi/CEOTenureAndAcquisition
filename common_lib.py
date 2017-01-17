from operator import itemgetter
import numpy as np
from itertools import groupby
from collections import OrderedDict


def Binary_Search_dictList(data_table, StockId, Year, GM_Name):
    lenT = int(len(data_table))
    middle = lenT // 2
    if lenT < 50000:
        foundItem = [x for x in data_table[:middle] if x['Stkcd'] == StockId \
                     and str(x['Reptdt']).split('-')[0] == Year \
                     and x['D0101b'] == GM_Name]
        if len(foundItem) > 0:
            return foundItem
        else:
            return []
    foundItem = [x for x in data_table[:middle] if x['Stkcd'] == StockId]
    # and str(x['Reptdt']).split('-')[0] == Year \
    # and x['D0101b'] == GM_Name]
    if len(foundItem) > 0:
        foundItem = [x for x in foundItem if str(x['Reptdt']).split('-')[0] == Year and x['D0101b'] == GM_Name]
        if len(foundItem) > 0:
            return foundItem
        else:
            return []
    else:
        foundItem = list(Binary_Search_dictList(list(data_table[middle:]), StockId, Year, GM_Name))
        if len(foundItem) > 0:
            return foundItem
        else:
            return []


def Group_Search_dictList(data_table, StockId, Year, GM_Name):
    for Stkcd,items in groupby(data_table, key=itemgetter('Stkcd')):
        if Stkcd==StockId:
            tempList = []
            for i in items:
                tempList.append(OrderedDict(i))
            foundItem=[x for x in tempList if x['Stkcd'] == StockId \
                     and str(x['Reptdt']).split('-')[0] == Year \
                     and x['D0101b'] == GM_Name]
            if len(foundItem)>0:
                return foundItem
            else:
                return []


def filter_dict_list_equal_or_not(data_table, filter_key, filter_value, isEqual):
    if isEqual:
        result = list(filter(lambda item: item[filter_key] == filter_value, data_table))
    else:
        result = list(filter(lambda item: item[filter_key] != filter_value, data_table))
    return result


def filter_dict_list_contain_or_not(data_table, filter_key, filter_value, isContain):
    if isContain:
        result = list(filter(lambda item: filter_value in item[filter_key], data_table))
    else:
        result = list(filter(lambda item: filter_value not in item[filter_key], data_table))
    return result


def extract_dict_list_to_NumPyArray(table, key):
    return np.array([i for i in [j[key] for j in table]])


def sort_dict_list(data_table, sort_keys):
    '''
    排序字典列表，
    :param data_table: a dict list
    :param sort_keys: a list of keys to sort
    :return: a sorted dict list
    '''
    lenKeys = len(sort_keys)
    if lenKeys == 1:
        result = sorted(data_table, key=itemgetter(sort_keys[0]))
    elif lenKeys == 2:
        result = sort_dict_list_two_keys(data_table, sort_keys)
    elif lenKeys == 3:
        result = sort_dict_list_three_keys(data_table, sort_keys)
    return result


def sort_dict_list_two_keys(data_table, sort_keys):
    return sorted(data_table, key=itemgetter(sort_keys[0], sort_keys[1]))


def sort_dict_list_three_keys(data_table, sort_keys):
    return sorted(data_table, key=itemgetter(sort_keys[0], sort_keys[1], sort_keys[2]))


def write_dict_list_to_TxtFile(data_table):
    pass


def print_dict_list(data_table):
    for item in data_table:
        print(item)
