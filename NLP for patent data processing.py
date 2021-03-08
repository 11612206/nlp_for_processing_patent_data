# import packages
import xlwt
import xlrd
import jieba
import string
import json
import nltk
import os

# pre-configuration for English tokenisation
mwe = [('United', 'States'),('United', 'States','Of','America'),('New','Hampshire'),('New', 'Jersey'),('New',
'Mexico'),
('New','York'),('North', 'Carolina'),('North', 'Dakota'),('Rhode', 'Island'),('South', 'Carolina'),
('South', 'Dakota'),('West','Virginia'),('District', 'Of','Columbia'),('United','Kingdom'),
('Northern','Ireland'),
('British','Columbia'), ('New','Brunswick'), ('Newfoundland','And','Labrador'),('Northwest','Territories'),
('Nova','Scotia'),('Prince','Edward','Island'),('Hong','Kong'),('Inner','Mongolia'),('Guang','Dong')]
mwe_tokenizer = nltk.tokenize.MWETokenizer(mwe) # add mwe

# define the function for identifying Chinese
def is_contain_chinese(check_str):
    for ch in check_str:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
        return False

# read location dictionary
file = open('dictionary.txt', 'r')
js = file.read()
diction_region = json.loads(js)
file.close()
file2 = open('dictionary_firm.txt', 'r')
js2 = file2.read()
diction_firm = json.loads(js2)
file2.close()
new_diction_firm = {v: k for k, v in diction_firm.items()} #反转字典
# read company-code document
file3 = open('dictionary_code.txt', 'r')
js3 = file3.read()
diction_code = json.loads(js3)
file3.close()
# read Chinese location dictionary
file4 = open('dictionary2.txt', 'r')
js4 = file4.read()
diction_CN = json.loads(js4)
file4.close()
# read patent data 
filelist=[]
sheet_list=[]
for root, dirs, files in os.walk(".", topdown=False):
    for name in files:
        char=os.path.join(root, name)
        if char.split('.')[-1]=='XLS':
            filelist.append(char)
for i in filelist: # navigate each file
    book = xlrd.open_workbook(i) # read workbook
    temp_list=book.sheet_names()
    for j in temp_list:
        sheet = book.sheet_by_name(j)
        sheet_list.append(sheet) #把所有文件夹内所有工作表全部储存至列表里面
# get # of companies
firm_num = len(diction_firm)
# construct year dictionary
diction_year={2016:-1,2017:0}
# result list
A=[[0 for i in range(16)]for i in range(2*firm_num)] #需要随着地区的添加随着修改！！！！！
row=0
column=0
exist=0
# for loop to do tokenisation
for sheet_item in sheet_list:
    nrows = sheet_item.nrows
    ncols = sheet_item.ncols
    for nrow in range(1,nrows):
        if sheet_item.cell(nrow, 0).value: 
            firm_all=sheet_item.cell(nrow, 29).value
            for item in firm_all.split(' | '):
                if item in diction_firm:
                    exist=1
                    code=(diction_firm[item]-1)*2
                    plus=diction_year[int(sheet_item.cell(nrow, 7).value)]
                    row=code+plus+1
                    A[row][15]+=1
                    break
                else:
                    exist=0
    elif exist!=0: # patent citation
        A[row][2]+=1
        string = sheet_item.cell(nrow, 32).value.title()
        flag=0
        if is_contain_chinese(string): # use jieba to tokenise
            list_address = list(jieba.cut(string)) # store each address
            for w in list_address:
                if w in diction_region:
                    column=diction_region[w]+4
                    A[row][column] += 1
                    flag=1
                    break
            if flag==0: # select Chinese patent
                for w in list_address:
                    if w in diction_CN:
                        column = 3 
                        A[row][column] += 1
                        break
        else: # if address doesn't contain Chinese,use nltk to process
            tokenized_string = nltk.word_tokenize(string) # pre-processing
            result_item = mwe_tokenizer.tokenize(tokenized_string)
            # print(result_item)
            lresult_item=list(result_item)
            lresult_item.reverse() 
            for w in lresult_item:
                if w in diction_region:
                    column = diction_region[w] +4 
                    A[row][column] += 1
                    flag=1
                    break
            if flag==0:
                for w in lresult_item:
                    if w in diction_CN:
                        column = 3
                        A[row][column] += 1
                        break

# print out the result
for i in range(len(A)): #i row
    for j in range(len(A[0])): #j column
        if j==0: # year
            A[i][j]=2016+i%2
            print(A[i][j], end=' ')
        elif j==1: # company code
            a=new_diction_firm[i//2+1]
            # matching
            A[i][j]=diction_code[a]
            print(A[i][j], end=' ')
        elif j==4:
            A[i][j]=A[i][2]-A[i][3]
            print(A[i][j], end=' ')
        elif j == len(A[0]) - 1:
            print(A[i][j])
        elif j == 14:
            A[i][j] = A[i][4] - A[i][5] - A[i][6] - A[i][7] - A[i][8] - A[i][9] - A[i][10] - A[i][11] - A[i][12] - A[i][13]
            print(A[i][j], end=' ')
        else:
            print(A[i][j], end=' ')