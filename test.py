from operator import itemgetter
import common_lib
import numpy as np
from sklearn.linear_model import LinearRegression

values = [{'Position': 2.0, 'Dimreas': '', 'Name': '胡跃飞', 'Chgdt': '2016-12-10', 'Changtyp': 2.0, 'ACME': 45.23},
          {'Position': 2.0, 'Dimreas': '', 'Name': '胡跃飞', 'Chgdt': '2016-12-10', 'Changtyp': 2.0, 'ACME': 612.78},
          {'Position': 2.0, 'Dimreas': 12.0, 'Name': '胡跃飞', 'Chgdt': '2016-12-12', 'Changtyp': 1.0, 'ACME': 205.55},
          {'Position': 1.0, 'Dimreas': '', 'Name': '谢永林', 'Chgdt': '2016-12-10', 'Changtyp': 2.0, 'ACME': 37.20},
          {'Position': 1.0, 'Dimreas': 12.0, 'Name': '姚波', 'Chgdt': '2016-12-10', 'Changtyp': 1.0, 'ACME': 10.75},
          {'Position': 1.0, 'Dimreas': '', 'Name': '姚波', 'Chgdt': '2016-11-07', 'Changtyp': 2.0, 'ACME': 55.23}]


def is_int(item):
    if item['Position'] == 1.0:
        return True
    return False


def count(n):
    while n > 0:
        yield n
        n -= 1


# ivals = list(filter(is_int, values))
# ivals = list(filter(lambda item: item['Position'] == 1.0, values))

# for item in ivals:
#     print(item)

# sort_keys = ('Position', 'Name')
# print(sort_keys)
# result = sorted(values, key=itemgetter(sort_keys[0], sort_keys[1]))
# common_lib.print_dict_list(result)

a = np.array([j for j in [[i['ACME']] for i in values]])
print(a)
print(type(a))
print(a.sum(axis=0))
print()


# b = common_lib.dedupe(values, key=lambda d: (d['Position'], d['Changtyp']))
# common_lib.print_dict_list(b)

# c = count(5)
# print(type(c))
# print(c)
# # print(c.__next__())
# # print(c.__next__())
# b = list(c)
# print(b)
