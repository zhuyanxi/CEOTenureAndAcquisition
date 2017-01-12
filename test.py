from operator import itemgetter
import common_lib

values = [{'Position': 2.0, 'Dimreas': '', 'Name': '胡跃飞', 'Chgdt': '2016-12-10', 'Changtyp': 2.0},
          {'Position': 2.0, 'Dimreas': '', 'Name': '胡跃飞', 'Chgdt': '2016-12-10', 'Changtyp': 2.0},
          {'Position': 2.0, 'Dimreas': 12.0, 'Name': '胡跃飞', 'Chgdt': '2016-12-10', 'Changtyp': 1.0},
          {'Position': 1.0, 'Dimreas': '', 'Name': '谢永林', 'Chgdt': '2016-12-10', 'Changtyp': 2.0},
          {'Position': 1.0, 'Dimreas': 12.0, 'Name': '姚波', 'Chgdt': '2016-12-10', 'Changtyp': 1.0},
          {'Position': 1.0, 'Dimreas': '', 'Name': '姚波', 'Chgdt': '2016-11-07', 'Changtyp': 2.0}]


def is_int(item):
    if item['Position'] == 1.0:
        return True
    return False


# ivals = list(filter(is_int, values))
# ivals = list(filter(lambda item: item['Position'] == 1.0, values))

# for item in ivals:
#     print(item)

# sort_keys = ('Position', 'Name')
# print(sort_keys)
# result = sorted(values, key=itemgetter(sort_keys[0], sort_keys[1]))
# common_lib.print_dict_list(result)

s=set(values)
