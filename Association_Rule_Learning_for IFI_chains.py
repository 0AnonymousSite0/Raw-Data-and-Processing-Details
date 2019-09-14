#-*- coding: utf-8 -*-

from __future__ import print_function
import pandas as pd
from apriori import *

inputfile = 'C:/Users/zhou/Desktop/rrecognized affected infrastructure entities in each news set.xls'
outputfile = 'C:/Users/zhou/Desktop/apriori_rules.xls'
data = pd.read_excel(inputfile, header = None)

print(u'\nstart')
ct = lambda x : pd.Series(1, index = x[pd.notnull(x)])
b = map(ct, data.as_matrix())
data = pd.DataFrame(list(b)).fillna(0)
print(u'\nfinish')
del b

support = 0.00
confidence = 0.00
ms = '---'

find_rule(data, support, confidence, ms).to_excel(outputfile)