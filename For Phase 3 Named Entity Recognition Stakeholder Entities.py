#encoding=utf-8
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

workbook = xlsxwriter.Workbook('C:\\Users\\zhou\\Desktop\\interdependence\\InterdependentStakeholders2020.xlsx')
WhetherTheRelationshipIsExistedWorksheet = workbook.add_worksheet('WhetherTheRelationshipIsExisted')
RMWorksheet = workbook.add_worksheet('RelationshipMatrix')
RMWorksheetAll = workbook.add_worksheet('RelationshipMatrixAll')
ColumnofText =0
Tri2Sim = OpenCC('t2s')
# def Tri2Sim(line):
#     sentence = Converter('zh-hans').convert(line)
#     sentence.encode('utf-8')
#     return sentence
myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
NewsData = xlrd.open_workbook("C:\\Users\\zhou\\Desktop\\news.xlsx")  # source of news
TableNewsTexts = NewsData.sheet_by_name(u'revisedcon')  # establish self-dic
jieba.load_userdict("C:\\Users\\zhou\\Desktop\\interdependence\\dic.txt")  # establ0ish self-dic
NumberOfNews = TableNewsTexts.nrows
NewsIndex = 0  # No. of the news
TargetEntitiesData = xlrd.open_workbook('C:\\Users\\zhou\\Desktop\\interdependence\\TargtEntitiesData.xlsx')
TableEntitiesData = TargetEntitiesData.sheet_by_name(u'Sheet1')
#TableEntitiesData = TargetEntitiesData.sheet_by_name(u'damage')
TargetEntities = TableEntitiesData.col_values(0)
def stopwordslist(filepath):
    stopwords=[]
    for line in open(filepath, 'r', encoding='utf-8'):
        line = line.encode('utf-8').decode('utf-8-sig')
        line = line.strip('\n')
        stopwords.append(line)

    return stopwords
stopwords = stopwordslist('C:\\Users\\zhou\\Desktop\\interdependence\\HITstopwords.txt')
# print (stopwords)
combine_dict = {}

for line in open("C:\\Users\\zhou\\Desktop\\interdependence\\Synonyms.txt", "r", encoding='utf-8'):
#for line in open("C:\\Users\\zhou\\Desktop\\interdependence\\NoneS.txt", "r", encoding='utf-8'):

    line = line.encode('utf-8').decode('utf-8-sig')  # very important and solve the problem of "\ufeff"
    seperate_word = line.strip().split(" ")
    num = len(seperate_word)
    for i in range(1, num):
        combine_dict[seperate_word[i]] = seperate_word[0]

WhetherTheRelationshipIsExisted = np.zeros((NumberOfNews, len(TargetEntities)), dtype=np.int)

for NewsItem in TableNewsTexts.col_values(ColumnofText):
    NewsItem = Tri2Sim.convert(NewsItem)

    seg_list = jieba.cut(NewsItem, cut_all=False

                         )
    #seg_list = jieba.cut(NewsItem)

    # Remove stopwords
    Result = ''
    for word in seg_list:
        if word not in stopwords:
            if word != '\t':
                if len(word)!=0:
                    Result += word
                    Result += " "

    ResultWithoutStopWords = Result.split(" ")
    #print (ResultWithoutStopWords)
    if  ResultWithoutStopWords==None:
        c = open("C:\\Users\\zhou\\Desktop\\interdependence\\word_vector_code_3.txt", 'a')
        c.write('\n')
    else:
        a=[]
        for i in ResultWithoutStopWords:
            a.append(i)
            #print(type(ResultWithoutStopWords))
            c = open("C:\\Users\\zhou\\Desktop\\interdependence\\word_vector_code_3.txt", 'a',encoding ='utf-8-sig' )
         #   i.encode('utf-8').decode('utf-8-sig')



        for i,p in enumerate(ResultWithoutStopWords):
            c.write(str(i+1)+' '+p+'|')
            #c.write('|'.join(ResultWithoutStopWords))
        c.write('\n')
   # c.close()


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

    array_index_of_word=[]
    array_word=[]
    single_array_word=[]
    for index_of_word, Token in enumerate(ResultWithoutStopWordsAndSynonyms):
        # print(Token[0])
        for index in range(len(TargetEntities)):
            if Token == TargetEntities[index]:
                array_index_of_word.append(index_of_word+1)
                array_word.append(Token)
                #print(index_of_word+1)
                WhetherTheRelationshipIsExisted[NewsIndex][index] = 1
    # print (NewsIndex)
    if  array_index_of_word==None:
        f = open("C:\\Users\\zhou\\Desktop\\interdependence\\index_stake.txt", 'a')
        f.write('\n')
    else:
        for i in range(len(array_index_of_word)):
            f = open("C:\\Users\\zhou\\Desktop\\interdependence\\index_stake.txt", 'a',encoding='utf-8-sig')
            f.write(str(array_index_of_word[i])+' '+str(array_word[i])+',')
        f.write('\n')
 #   f.close()

    NewsIndex = NewsIndex + 1

    #print (ResultWithoutStopWordsAndSynonyms)


# print (WhetherTheRelationshipIsExisted)
NumberOfRowWhetherTheRelationshipIsExisted = 0
NumberOfColumnWhetherTheRelationshipIsExisted = 0
for NumberOfRowWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted)):
    for NumberOfColumnWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted[0])):
        WhetherTheRelationshipIsExistedWorksheet.write_string(NumberOfRowWhetherTheRelationshipIsExisted,
                                                       NumberOfColumnWhetherTheRelationshipIsExisted, str( WhetherTheRelationshipIsExisted[NumberOfRowWhetherTheRelationshipIsExisted][NumberOfColumnWhetherTheRelationshipIsExisted]))
        # worksheet.write(NumberofRow, NumberofColumn, WhetherTheRelationshipIsExisted[m][index])
# workbook.save('C:\\Users\\zhou\\Desktop\\interdependence\\Matrix.xls')
# for tk in result:
#     print("word %s\t\t start: %d \t\t end:%d" % (tk[0], tk[1], tk[2]))

RM = np.zeros((len(TargetEntities), len(TargetEntities)), dtype=np.int)

for NumberOfRowWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted)):
    for NumberOfColumnWhetherTheRelationshipIsExisted in range(len(WhetherTheRelationshipIsExisted[0])):
        if WhetherTheRelationshipIsExisted[NumberOfRowWhetherTheRelationshipIsExisted][
            NumberOfColumnWhetherTheRelationshipIsExisted] == 1:
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