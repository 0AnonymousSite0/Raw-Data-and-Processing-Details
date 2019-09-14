# -*-coding:utf-8-*-
from opencc import OpenCC
import datetime
starttime = datetime.datetime.now()
print(starttime)
#long running
#do something other
import xlsxwriter
import jieba
import xlrd
# import xlwt
import jieba.analyse
import numpy as np
# import uniout
NumberOfNews=2731
Tri2Sim = OpenCC('t2s')
# hash操作
# hash操作
def string_hash(source):
    if source == "":
        return 0
    else:
        x = ord(source[0]) << 7
        m = 1000003
        mask = 2 ** 128 - 1
        for c in source:
            x = ((x * m) ^ ord(c)) & mask
        x ^= len(source)
        if x == -1:
            x = -2
        x = bin(x).replace('0b', '').zfill(64)[-64:]
        # print(source, x)  # 打印 （关键词，hash值）
        return str(x)

SimilarityMatrix = np.zeros((NumberOfNews, NumberOfNews), dtype=np.int)
n=1
for line in open('C:\\Users\\zhou\\Desktop\\interdependence\\allnews.txt', encoding='UTF-8').readlines():

    seg = jieba.cut(Tri2Sim.convert(line))  # 分词
    jieba.analyse.set_stop_words('C:\\Users\\zhou\\Desktop\\interdependence\\stopwords.txt')  # 去除停用词
    keyWord = jieba.analyse.extract_tags('|'.join(seg), topK=20, withWeight=True, allowPOS=())  # 先按照权重排序，再按照词排序
     # print (keyWord)  # 前20个关键词，权重
    keyList = []
    for feature, weight in keyWord:  # 对关键词进行hash
        weight = int(weight * 10)
        feature = string_hash(feature)
        temp = []
        for i in feature:
            if (i == '1'):
                temp.append(weight)
            else:
                temp.append(-weight)
        # print (temp)  # 将hash值用权值替代
        keyList.append(temp)
    list_sum = np.sum(np.array(keyList), axis=0)  # 20个权值列向相加
    # print 'list_sum:', list_sum  # 权值列向求和

    if (keyList == []):  # 编码读不出来
        print (n)
    simhash = ''

    for i in list_sum:  # 权值转换成hash值
        if (i > 0):
            simhash = simhash + '1'
        else:
            simhash = simhash + '0'
    simhash_1 = simhash
    print (simhash_1)  # str 类型
    n=n+1



