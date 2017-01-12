from operator import itemgetter


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


def print_dict_list(data_table):
    for item in data_table:
        print(item)
