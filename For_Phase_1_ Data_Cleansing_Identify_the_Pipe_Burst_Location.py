# encoding=utf-8
# -*- coding: utf8 -*-
import xlrd
import xlwt
import jieba
import numpy as np
import xlsxwriter
from opencc import OpenCC
import datetime
starttime = datetime.datetime.now()
print(starttime)

workbook = xlsxwriter.Workbook('****Excel for storing the results****')
WhetherTheRelationshipIsExistedWorksheet = workbook.add_worksheet('****Sheet for storing the results****')
RMWorksheet = workbook.add_worksheet('****Sheet for storing the results****')
RMWorksheetAll = workbook.add_worksheet('****Sheet for storing the results****')
ColumnofText = 0
Tri2Sim = OpenCC('t2s')

myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
NewsData = xlrd.open_workbook("****file path of raw news****")  # source of news
TableNewsTexts = NewsData.sheet_by_name"****sheet of the news****")
jieba.load_userdict("****file path of more than 5000 roads in Hong Kong****")  # It is part of the Hong Kong DIC
NumberOfNews = TableNewsTexts.nrows
NewsIndex = 0  # No. of the news
TargetEntitiesData = xlrd.open_workbook("****file path of more than 5000 roads in Hong Kong****")
TableEntitiesData = TargetEntitiesData.sheet_by_name(u'****Sheet name of more than 5000 roads in Hong Kong****')
TargetEntities = TableEntitiesData.col_values(0)
print(TargetEntities)
def stopwordslist(filepath):
    stopwords=[]
    for line in open(filepath, 'r', encoding='utf-8'):
        line = line.encode('utf-8').decode('utf-8-sig')
        line = line.strip('\n')
        stopwords.append(line)

    return stopwords
stopwords = stopwordslist("****file path of stop words****")
# print (stopwords)
combine_dict = {}

for line in open("provide a blank txt, as there is no synonyms of roads locations", "r", encoding='utf-8'):
    line = line.encode('utf-8').decode('utf-8-sig')  # very important and solve the problem of "\ufeff"
    seperate_word = line.strip().split(" ")
    num = len(seperate_word)
    for i in range(1, num):
        combine_dict[seperate_word[i]] = seperate_word[0]

WhetherTheRelationshipIsExisted = np.zeros((NumberOfNews, len(TargetEntities)), dtype=np.int)

for NewsItem in TableNewsTexts.col_values(ColumnofText):
    NewsItem = Tri2Sim.convert(NewsItem)

    seg_list = jieba.cut(NewsItem, cut_all=True)
    tokenize_list=jieba.tokenize(NewsItem,mode='search')

    # Remove stopwords
    Result = ''
    for word in seg_list:
        if word not in stopwords:
            if word != '\t':
                Result += word
                Result += " "
    ResultWithoutStopWords = Result.split(" ")
    # result = jieba.tokenize(NewsItem, mode='search')
    ResultWithoutStopWordsAndSynonyms = ''


    # Handle Synonyms
    for WordWithoutStopWords in ResultWithoutStopWords:
        if WordWithoutStopWords in combine_dict:
            WordWithoutStopWords = combine_dict[WordWithoutStopWords]
            ResultWithoutStopWordsAndSynonyms += WordWithoutStopWords
            ResultWithoutStopWordsAndSynonyms += " "
        else:
            ResultWithoutStopWordsAndSynonyms += WordWithoutStopWords
            ResultWithoutStopWordsAndSynonyms += " "
    ResultWithoutStopWordsAndSynonyms = ResultWithoutStopWordsAndSynonyms.split(" ")

    # Identify entities
    for Token in range(len(ResultWithoutStopWordsAndSynonyms)):
        # print(Token[0])
        for index in range(len (TargetEntities)):
            # print(Token)
            # print(TargetEntities[index])
            # print(Token == TargetEntities[index])
            if ResultWithoutStopWordsAndSynonyms[Token] == TargetEntities[index]:
                # WhetherTheRelationshipIsExisted[NewsIndex][index] = 1
                print("Newsno%s\t word %s\t position: %d \t" % (NewsIndex+1, TargetEntities[index], Token))
                # print (NewsIndex+1, TargetEntities[index], Token)
                # for tk in tokenize_list:
                #     if tk[0]==Token:
                #         print("Newsno%s\t\tword %s\t\t start: %d \t\t end:%d" % (NewsIndex, tk[0], tk[1], tk[2]))




    # print (NewsIndex)
    NewsIndex = NewsIndex + 1
    # print (tokenize_list)
    # print (ResultWithoutStopWordsAndSynonyms)

# print (WhetherTheRelationshipIsExisted)
NumberOfRowWhetherTheRelationshipIsExisted = 0
NumberOfColumnWhetherTheRelationshipIsExisted = 0
for NumberOfRowWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted)):
    for NumberOfColumnWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted[0])):
        WhetherTheRelationshipIsExistedWorksheet.write_string(NumberOfRowWhetherTheRelationshipIsExisted,NumberOfColumnWhetherTheRelationshipIsExisted, str( WhetherTheRelationshipIsExisted[NumberOfRowWhetherTheRelationshipIsExisted][NumberOfColumnWhetherTheRelationshipIsExisted]))
        # worksheet.write(NumberofRow, NumberofColumn, WhetherTheRelationshipIsExisted[m][index])
# workbook.save('C:\\Users\\zhou\\Desktop\\interdependence\\Matrix.xls')
# for tk in result:
#     print("word %s\t\t start: %d \t\t end:%d" % (tk[0], tk[1], tk[2]))

RM = np.zeros((len(TargetEntities), len(TargetEntities)), dtype=np.int)

for NumberOfRowWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted)):
    for NumberOfColumnWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted[0])):
        if WhetherTheRelationshipIsExisted[NumberOfRowWhetherTheRelationshipIsExisted][NumberOfColumnWhetherTheRelationshipIsExisted] == 1:
            # a= NumberofRowWhetherTheRelationshipIsExisted
            b = NumberOfColumnWhetherTheRelationshipIsExisted
            for a in range(b + 1, len(WhetherTheRelationshipIsExisted[0])):
                if WhetherTheRelationshipIsExisted[NumberOfRowWhetherTheRelationshipIsExisted][a] == 1:
                    RM[NumberOfColumnWhetherTheRelationshipIsExisted][a] = RM[NumberOfColumnWhetherTheRelationshipIsExisted][a] + 1
# print(RM)

RowofRM = 0
ColumnofRM = 0
for RowofRM in range(len(RM)):
    for ColumnofRM in range(len(RM[0])):
        RMWorksheet.write_string(RowofRM, ColumnofRM, str(RM[RowofRM][ColumnofRM]))
        RMWorksheetAll.write_string(ColumnofRM, RowofRM, str(RM[RowofRM][ColumnofRM]))
        # worksheet.write(NumberofRow, NumberofColumn, WhetherTheRelationshipIsExisted[m][index])
workbook.close()
endtime = datetime.datetime.now()
print(endtime)
print((endtime-starttime).seconds)
