import random
import xlrd
from openpyxl import load_workbook
from matplotlib import pyplot
from copy import deepcopy

file_location = "C:/Users/neel/Documents/10 Data Science competitions/02 Data Science Fellowship/Data/election/2015/P00000001-AZ.xlsx"
wb = xlrd.open_workbook(file_location)
sheet = wb.sheet_by_index(0)

## data = [combined_var,contbr_city,zip,employer,occupation, combined_amt]

data_intermediate = [[sheet.cell_value(r,c) for c in [x for x in range(sheet.ncols) if ((x != 0)&(x != 1)&(x != 5)& (x != 10)&(x != 11)&(x != 12)&(x != 13)&(x != 14)&(x != 15)&(x != 16)) ]] for r in range(sheet.nrows)]

data = []
data_intermediate1 = []

for inst in range(1,len(data_intermediate)):
    if (data_intermediate[inst][2]=="OTTAWA")or (data_intermediate[inst][2]=="LISTOWEL ONTARIO C")or \
       (data_intermediate[inst][2]=="HAMILTON")or(data_intermediate[inst][2]=="MCDONOUGH") \
        or (data_intermediate[inst][3]=='')or (data_intermediate[inst][2]=='')\
        or (data_intermediate[inst][4]=='')or (data_intermediate[inst][5]=='')\
        or (data_intermediate[inst][6]=='')or (data_intermediate[inst][7]=='')\
        or (data_intermediate[inst][2]=='')or (data_intermediate[inst][1]=='')\
        or (data_intermediate[inst][0]==''):
        continue
    else:
        data_intermediate1.append(deepcopy(data_intermediate[inst]))

import operator
def sort_table(table, col=0):
    return sorted(table, key=operator.itemgetter(col),reverse=True)
for row in sort_table(data_intermediate1, 6):
    data.append(deepcopy(row))

occupation = []
occupation_dic = {}
occupation_number = 1.0

city = []
city_dic = {}
city_number = 1.0

contbr_employer = []
contbr_employer_dic = {}
contbr_employer_number = 1.0

election_tp = []
election_tp_dic = {}
election_tp_number = 1.0

cand_nm_cont_nm_ele_tp_dic = {}
cand_nm_cont_nm_ele_tp_no_dic = {}
cand_nm_cont_nm_ele_tp_number = 1.0

cand_nm = []
cand_nm_dic = {}
cand_nm_number = 1.0

for instance in range(1,len(data)):
    if data[instance][5] not in occupation:
        occupation.append(data[instance][5])
    if data[instance][2] not in city:
        city.append(data[instance][2])
    if data[instance][4] not in contbr_employer:
        contbr_employer.append(data[instance][4])
    if data[instance][0] not in cand_nm:
        cand_nm.append(data[instance][0])
    
    local_var0 = data[instance][0]+data[instance][1]+data[instance][7]
    if local_var0 not in cand_nm_cont_nm_ele_tp_dic:
        cand_nm_cont_nm_ele_tp_dic[local_var0]=data[instance][6]
    if local_var0 in cand_nm_cont_nm_ele_tp_dic:
        cand_nm_cont_nm_ele_tp_dic[local_var0]=cand_nm_cont_nm_ele_tp_dic[local_var0]+data[instance][6]

    
## data = [cand_nm,contbr_nm,contbr_city,zip,employer,
## occupation, amt, election_tp]

for type_occ in occupation:
    occupation_dic[type_occ] = occupation_number
    occupation_number = occupation_number + 1.0
for type_city in city:
    city_dic[type_city] = city_number
    city_number = city_number + 1.0
for type_contbr_employer in contbr_employer:
    contbr_employer_dic[type_contbr_employer] = contbr_employer_number
    contbr_employer_number = contbr_employer_number + 1.0
for keys in cand_nm_cont_nm_ele_tp_dic:
    cand_nm_cont_nm_ele_tp_no_dic[keys] = cand_nm_cont_nm_ele_tp_number
    cand_nm_cont_nm_ele_tp_number = cand_nm_cont_nm_ele_tp_number + 1.0
for type_cand_nm in cand_nm:
    cand_nm_dic[type_cand_nm] = cand_nm_number
    cand_nm_number = cand_nm_number + 1.0
    
help_to_del = []

for instance in range(1,len(data)):
    if data[instance][5] in occupation_dic:
        local_var1 = data[instance][5]
        data[instance][5] = occupation_dic[local_var1]
    if data[instance][2] in city_dic:
        local_var2 = data[instance][2]
        data[instance][2] = city_dic[local_var2]
    if data[instance][4] in contbr_employer_dic:
        local_var3 = data[instance][4]
        data[instance][4] = contbr_employer_dic[local_var3]
    
    local_var5 = data[instance][0]+data[instance][1]+data[instance][7]
    data[instance][1]=cand_nm_cont_nm_ele_tp_no_dic[local_var5]
    data[instance][6]=cand_nm_cont_nm_ele_tp_dic[local_var5]

    if data[instance][0] in cand_nm_dic:
        local_var6 = data[instance][0]
        data[instance][0] = cand_nm_dic[local_var6]

## data = [cand_nm,combined_var,contbr_city,zip,employer,
## occupation, combined_amt, election_tp]

final_data = []

for instance in range(1,len(data)):
    if data[instance] not in final_data:
        final_data.append(data[instance][:-1])
    else:
        continue
## data = [cand_nm,combined_var,contbr_city,zip,employer,
## occupation, combined_amt]

##end of data clean

##start of split in train and test data
len_fin = len(final_data)
total_list = random.sample(range(0, len_fin), len_fin)

train_att = []
train_class = []
test_att = []
test_class = []
for i in range (0,len_fin*4/5):
    j = total_list[i]
    train_att.append(final_data[j][1:])
    train_class.append(final_data[j][0])
    
for i in range (len_fin*4/5,len_fin):
    j = total_list[i]
    test_att.append(final_data[j][1:])
    test_class.append(final_data[j][0])

from sklearn.ensemble import RandomForestClassifier
rf =RandomForestClassifier(n_estimators=30,criterion='entropy',max_features=3, \
                             min_samples_leaf=5,bootstrap=True)
print('fitting the model')
rf.fit(train_att, train_class)
predicted_class = rf.predict(test_att)
importances = rf.feature_importances_
    
cand_nm_rev_dic = {}
for key in cand_nm_dic:
    cand_nm_rev_dic[cand_nm_dic[key]] = key


from collections import defaultdict
d = defaultdict(int)
for i in predicted_class:
    d[i] += 1
result = max(d.iteritems(), key=lambda x: x[1])

print "the winner is" , cand_nm_rev_dic[result[0]]
print importances

##text_file = open("employer.txt", "w")
##text_file.write("%s" % contbr_employer_dic)
##text_file.close()
##text_file = open("occupation.txt", "w")
##text_file.write("%s" % occupation_dic)
##text_file.close()



