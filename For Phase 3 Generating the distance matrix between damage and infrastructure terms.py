
import xlrd
import xlwt
import jieba
import numpy as np
import xlsxwriter
from opencc import OpenCC
import datetime
starttime = datetime.datetime.now()
print(starttime)
TargetEntitiesData = xlrd.open_workbook('C:\\Users\\zhou\\Desktop\\interdependence\\TargtEntitiesData.xlsx')
TableEntitiesData = TargetEntitiesData.sheet_by_name(u'infrastructure')
#TableEntitiesData = TargetEntitiesData.sheet_by_name(u'damage')
TargetEntities = TableEntitiesData.col_values(0)
RM = np.zeros((3000, len(TargetEntities)), dtype=np.int)
m=0
with open('C:\\Users\\zhou\\Desktop\\interdependence\\index_infra.txt','r', encoding='utf-8') as f1:
    with open('C:\\Users\\zhou\\Desktop\\interdependence\\index_da.txt','r',  encoding='utf-8') as f2:
        for x,y in zip(f1.readlines(),f2.readlines()):

#for line in open("C:\\Users\\zhou\\Desktop\\interdependence\\index_infra.txt", "r", encoding='utf-8'):
#for line in open("C:\\Users\\zhou\\Desktop\\interdependence\\NoneS.txt", "r", encoding='utf-8'):
            number=[]
            word=[]
            x = x.encode('utf-8').decode('utf-8-sig')  # very important and solve the problem of "\ufeff"

            seperate_word = x.strip().replace(' ',',').split(",")
            for i in range(len(seperate_word)-1):
                if i%2==0:
                    number.append(seperate_word[i])
                else:
                    word.append(seperate_word[i])
            number_da=[]
            word_da=[]
            y = y.encode('utf-8').decode('utf-8-sig')  # very important and solve the problem of "\ufeff"

            seperate_word = y.strip().replace(' ',',').split(",")
            for i in range(len(seperate_word)-1):
                if i%2==0:
                    number_da.append(seperate_word[i])
                else:
                    word_da.append(seperate_word[i])
            least_distance = [10000 for l in range(len(TargetEntities))]
            for n,p in enumerate(TargetEntities):



                for i,q in enumerate(word):

                    if p == q:
                        for s in number_da:


                            try:
                                current_distance=abs(int(number[i]) - int(s))
                                #print(current_distance)
                                if current_distance<least_distance[n]:
                                    #print(n)
                                    least_distance[n]=current_distance
                            except Exception as e:
                                print(e)
            #print(least_distance)
            for k,w in enumerate(least_distance):
                RM[m][k]=w
            m=m+1
            print(RM)
workbook = xlsxwriter.Workbook('C:\\Users\\zhou\\Desktop\\interdependence\\distance_between_Infra_da.xlsx')
RMWorksheet = workbook.add_worksheet('RelationshipMatrix')
RowofRM = 0
ColumnofRM = 0
for RowofRM in range(len(RM)):
    for ColumnofRM in range(len(RM[0])):
        RMWorksheet.write_string(RowofRM, ColumnofRM, str(RM[RowofRM][ColumnofRM]))
workbook.close()
        #   print (number)
          #  print(word)
           # print(seperate_word)

