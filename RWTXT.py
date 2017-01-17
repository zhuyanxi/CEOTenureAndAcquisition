import time
import read_excel as re
from operator import itemgetter
from collections import OrderedDict
import traceback

def write_txt():
    start = time.clock()
    print(start)
    tables = []
    for i in range(1, 10):
        # table = re.excel_table_byname(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
        table = re.excel_table_to_OrderedDict_bySheetName(file='CG_Director_ALL.xls', by_name=u'CEO' + str(i))
        print(len(table))
        tables.extend(table)
    print('len of tables: %d' % len(tables))
    tables.sort(key=itemgetter('Stkcd'))
    end = time.clock()
    print(end)
    print('read excel Running time:%s Seconds' % ((end - start)))

    startX = time.clock()
    output = open('CG_Director_ALL.txt', 'w', encoding='utf-8')
    titles = ''
    titleList = list(tables[0].keys())
    titles = '|'.join(titleList)
    output.writelines(titles)
    for item in tables:
        row = list(OrderedDict(item).values())
        print(row)
        row = [str(i) if type(i) == float else i for i in row]
        # print(row)
        data = '|'.join(row)
        output.write('\n')
        output.write(data)
    output.close()
    endX = time.clock()
    print('write txt Running time:%s Seconds' % ((endX - startX)))



def txt_to_dict_list(file):
    data = open(file, 'r', encoding='utf-8')
    lines=data.readlines()
    # return lines

    print(lines[0])
    print(lines[1])
    result=[]
    keys=str(lines[0][:-1]).split('|')
    print(keys)
    for i in range(1,len(lines)):
        row = str(lines[i][:-1]).split('|')
        # print(row)
        app = OrderedDict()
        for j in range(len(keys)):
            app[keys[j]] = row[j]
        result.append(app)
    return result


# if __name__ == "__main__":
#     startX = time.clock()
#     data = txt_to_dict_list('CG_Director_ALL.txt')
#     print(len(data))
#     endX = time.clock()
#     print('write txt Running time:%s Seconds' % ((endX - startX)))
#     print(data[0])
#     print(data[0]['Stkcd'])