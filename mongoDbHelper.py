import pymongo
from pymongo.mongo_client import MongoClient
import pymongo.errors
import traceback

from collections import OrderedDict
from operator import itemgetter
import time
import read_excel as re


# try:
#     # start = time.clock()
#     # print(start)
#     # tables = []
#     # for i in range(1, 10):
#     #     # table = re.excel_table_byname(file='CG_Director_new.xls', by_name=u'CEO' + str(i))
#     #     table = re.excel_table_to_OrderedDict_bySheetName(file='CG_Director_ALL.xls', by_name=u'CEO' + str(i))
#     #     print(len(table))
#     #     tables.extend(table)
#     # print('len of tables: %d' % len(tables))
#     # tables.sort(key=itemgetter('Stkcd'))
#     # end = time.clock()
#     # print(end)
#     # print('read excel Running time:%s Seconds' % ((end - start)))
#
#
#     mongoClient = MongoClient('localhost', 27017)
#     mongoDatabase = mongoClient.CEOAgeResearch
#     print("connect database successfully")
#     mongoCollection = mongoDatabase.CG_Director_ALL
#
#     # start = time.clock()
#     # print(start)
#     # mongoCollection.insert_many(tables)
#     # end = time.clock()
#     # print(end)
#     # print('insert into db Running time:%s Seconds' % ((end - start)))
#
#     start = time.clock()
#     print(start)
#     r_collection=mongoCollection.find({'Stkcd': '300139', 'Reptdt': '2014-12-31', 'D0101b': '崔劲'})[0]
#     # for row in r_collection:#{'Stkcd': '000001', 'Reptdt': '2015-12-31'}
#     #     # print("-" * 50)
#     #     print(row)
#     # rList=list(r_collection)
#     end = time.clock()
#     print(end)
#     print('read db Running time:%s Seconds' % ((end - start)))
# except pymongo.errors.PyMongoError as e:
#     print("mongo error: ", e)
#     traceback.print_exc()


def connect_to_db():
    try:
        mongoClient = MongoClient('localhost', 27017)
        return mongoClient
    except pymongo.errors.PyMongoError as e:
        print("mongo error: ", e)
        traceback.print_exc()


def Get_CG_Director_ALL(dbClient):
    try:
        db = dbClient.CEOAgeResearch
        db_collection = db.CG_Director_ALL
        return db_collection
    except pymongo.errors.PyMongoError as e:
        print("mongo error: ", e)
        traceback.print_exc()


def Search_CG_Director_ALL(db_collection, StockId, Year, GM_Name):
    r_collection = list(db_collection.find({'Stkcd': StockId, 'Reptdt': Year + '-12-31', 'D0101b': GM_Name}))
    if len(r_collection) > 0:
        return r_collection
    else:
        return []


# if __name__ == "__main__":
#     # r_collection = Search_CG_Director_ALL(Get_CG_Director_ALL(connect_to_db()),'300139','2014','崔劲')
#     docCollection = Get_CG_Director_ALL(connect_to_db())
#     print(type(docCollection))
#
#     startT = time.clock()
#     lenT = 0
#     for i in range(30000):
#         startX = time.clock()
#         # r_collection = db.find({'D0401b': {"$gte": 30, "$lte": 50}})
#         r_collection = docCollection.find({'Stkcd': '300139', 'Reptdt': '2014-12-31', 'D0101b': '崔劲'})
#         # r_collection = db.find([{'Stkcd': '300139'},{'Stkcd': '000002'}])
#         # app=OrderedDict()
#         # for row in r_collection:
#         # app['Stkcd']=row['Stkcd']
#         # app['Reptdt'] = row['Reptdt']
#         # app['Name'] = row['D0101b']
#         # print(row)
#         # len=0
#         # r_Iter=r_collection.__iter__()
#         # result=[i for i in r_collection]
#         # print(result)
#         print(r_collection[0])
#         print(type(r_collection))
#         endX = time.clock()
#         lenT += 1
#         print('%s Length' % lenT)
#         print('read one record Running time:%s Seconds' % (endX - startX))
#     # print(type(r_collection))
#     # for i in r_collection:
#     #     print(r_collection)
#     # print(len(list(r_collection)))
#     endT = time.clock()
#     print(endT)
#     print('Running time:%s Seconds' % ((endT - startT)))
