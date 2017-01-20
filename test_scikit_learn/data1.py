import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import numpy as np


def runplt():
    plt.figure()
    plt.title(u'diameter-cost curver')
    plt.xlabel(u'diameter')
    plt.ylabel(u'cost')
    plt.axis([0, 25, 0, 25])
    plt.grid(True)
    return plt


# plt = runplt()
X = [[6], [8], [10], [14], [18]]
y = [[7], [9], [13], [17.5], [18]]
print(type(X))
# plt.plot(X, y, 'k.')
# plt.show()

# 创建并拟合模型
model = LinearRegression()
model.fit(X, y)
print('预测一张12英寸匹萨价格：$%.2f' % model.predict(np.array([12]).reshape(-1, 1))[0])


plt = runplt()
plt.plot(X, y, 'k.')
X2 = [[0], [10], [14], [25]]
model = LinearRegression()
model.fit(X, y)
y2 = model.predict(X2)
plt.plot(X, y, 'k.')
plt.plot(X2, y2, 'g-')
plt.show()



# X = [[6, 2], [8, 1], [10, 0], [14, 2], [18, 0]]
# y = [[7], [9], [13], [17.5], [18]]
# model = LinearRegression()
# model.fit(X, y)
# X_test = [[8, 2], [9, 0], [11, 2], [16, 2], [12, 0]]
# y_test = [[11], [8.5], [15], [18], [11]]
# predictions = model.predict(X_test)
# for i, prediction in enumerate(predictions):
#     print('Predicted: %s, Target: %s' % (prediction, y_test[i]))
# print('R-squared: %.2f' % model.score(X_test, y_test))
