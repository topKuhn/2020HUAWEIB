import numpy as np
import math

def filterData(rowData):
    print("input size", rowData.shape)
    #存储含有粗大误差值的坏值的样本
    badRow = []
    for i in range(rowData.shape[1]):
        sum = 0
        for j in range(rowData.shape[0]):
            sum = sum + rowData[j][i]
        #第i列的平均值
        average = sum/rowData.shape[0]
        print("列" + str(i) + "平均值 ：", average)
        squareSum = 0
        for j in range(rowData.shape[0]):
            diffence = rowData[j][i] - average
            squareSum = squareSum + math.pow(diffence, 2)
        print("列" + str(i) + "平方和 ：", squareSum)
        σ = math.pow((squareSum/(rowData.shape[0] - 1)), 0.5)
        print("σ值：" , σ)
        for j in range(rowData.shape[0]):
            if(abs(rowData[j][i] - average) > (σ*3) ):
                if(j not in badRow):
                    badRow.append(j)


    result = []
    for i in range(rowData.shape[0]):
        if(i not in badRow):
            result.append(rowData[i])
    return badRow, result



a = [
    [1,2,4],
    [3,4,5],
    [9,2,2]
]
a = np.array(a)
badRow, result = filterData(a)
print(badRow)
print(result)