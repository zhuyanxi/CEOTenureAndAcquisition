import read_excel as re
from collections import OrderedDict
import common_lib
import numpy as np
import random
import math
from sklearn.linear_model import LinearRegression


def filter_data():
    tables = re.excel_table_to_OrderedDict_bySheetName(file='GM_Tenure_with_Param(WithCeoAge_VFinal2).xls',
                                                       by_name=u'data')
    tables = [i for i in tables if i['CashRatio'] != '' and i['GM_Name'] != '']
    return tables


if __name__ == "__main__":
    data_all = filter_data()
    common_lib.print_dict_list(data_all)
    print(len(data_all))

    training_data = random.sample(data_all, len(data_all) * 9 // 10)
    test_data = [i for i in data_all if i not in training_data]
    print(len(training_data))
    print(len(test_data))
    print(type(training_data[0]['LogTotalAsset(12-31)']))

    yVector = np.array([j for j in [[i['MATimes']] for i in training_data]])
    xVector = np.array([j for j in [[i['tenure'], i['LogTotalAsset(12-31)'],
                                     i['Lev'], i['FirmAge'], i['CEOAge']]
                                    for i in training_data]])
    print('-' * 50)
    print(yVector)
    print(xVector)

    yNormVector, yRanges, yMinVals = common_lib.auto_normalization(yVector)
    print('-' * 50)
    print(yNormVector)
    print(yRanges)
    print(yMinVals)

    xNormVector, xRanges, xMinVals = common_lib.auto_normalization(xVector)
    print('-' * 50)
    print(xNormVector)
    print(xRanges)
    print(xMinVals)

    # y_test = np.array([j for j in [[i['MATimes']] for i in test_data]])
    # x_test = np.array([j for j in [[i['tenure'], i['LogTotalAsset(12-31)'], i['Lev'], i['FirmAge'], i['CEOAge']]
    #                                for i in test_data]])
    # # print(y_test)
    # # print(x_test)
    #
    # model = LinearRegression()
    # model.fit(xVector, yVector)
    # print(model)
    # print(model.coef_)
    # print(model.intercept_)
    #
    # predictions = model.predict(x_test)
    # for i, prediction in enumerate(predictions):
    #     print('Predicted: %s, Target: %s' % (prediction, y_test[i]))
    # print('R-squared: %.2f' % model.score(x_test, y_test))
