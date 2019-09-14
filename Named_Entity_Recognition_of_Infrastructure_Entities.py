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
TableNewsTexts = NewsData.sheet_by_name("****sheet of the news****")  # establish self-dic
jieba.load_userdict("****file path of domain knowledge components****")  # establish self-dic
NumberOfNews = TableNewsTexts.nrows
NewsIndex = 0  # No. of the news
TargetEntitiesData = xlrd.open_workbook("****file path of potentially affected infrastructure entities****")
TableEntitiesData = TargetEntitiesData.sheet_by_name(u'****sheet of potentially affected infrastructure entities****')
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

combine_dict = {}

for line in open("****file path of synonyms of potentially affected infrastructure entities****", "r", encoding='utf-8'):
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
    # print(ResultWithoutStopWordsAndSynonyms)
    # Identify entities
    for Token in range(len(ResultWithoutStopWordsAndSynonyms)):
        # print(Token[0])
        for index in range(len (TargetEntities)):
            # print(Token)
            # print(TargetEntities[index])
            # print(Token == TargetEntities[index])
            if ResultWithoutStopWordsAndSynonyms[Token] == TargetEntities[index]:
                # WhetherTheRelationshipIsExisted[NewsIndex][index] = 1
                for i in range(1,5):  # here 5 is the distance threshold
                    if Token-i>0 and Token+i<len(ResultWithoutStopWordsAndSynonyms):
                        if (("****infrastructure damage-related words****"==ResultWithoutStopWordsAndSynonyms[Token-i]) or ("****infrastructure damage-related words****"==ResultWithoutStopWordsAndSynonyms[Token+i])) :
                            print( "Newsno%s\t word %s\t position: %d \t" % (NewsIndex + 1, TargetEntities[index], Token))
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