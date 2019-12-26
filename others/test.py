# # Suppose this is foo.py.

# print("before import")
# import math

# print("before functionA")
# def functionA():
#     print("Function A")

# print("before functionB")
# def functionB():
#     print("Function B {}".format(math.sqrt(100)))

# print("before __name__ guard")
# if __name__ == '__main__':
#     functionA()
#     functionB()
# print("after __name__ guard")

# li = ['qqq', '111', '1qw', '222']
# print(li)
# n_li=[]
# try:
#     n_li=list(map(int,li))
# except ValueError:

# for x in range(len(li)):
#     for y in range(len(li[x])):
#         n_li.append(li[x][y])
#         # if y == 1:

#         #     n_li.append(int(li[x][y]))


# li = [
#     ['1', '1', '2'],
#     ['1', '1', '2'],
#     ['1', '1', '2'],
#     ['1', '1', '2'],
#     ['1', '1', '5'],
#     ['1', '1', '5'],
#     ['1', '1', '8'],
#     ['1', '1', '8']
# ]


# insert_blank_line()
# print(n_li)
# ll = []
# li = ['坏账准备-应收账款', '存商品-半成品', '2', '3']
# error_name = ['2', '3']


# def replace_err_name(li,error_name):
#     for i in range(len(li)):
#         for x in range(len(error_name)):
#             if li[i] == error_name[x]:
#                 li[i]='xx'+li[i]

# replace_err_name()


# print(li)
# li = ['库存商品-半成品', '工程施工-差旅费-123123-444', '工程施工-行业会议', '工程施工-通信费']
# correct_name = '-'
# error_name = '工程施工-差旅费-123123-444'
# correct_name_taxx = '-'
# error_name_taxx = ['应交税费-应交增值税-111-不']


# def findStr(string, subStr, findCnt):
#     listStr = string.split(subStr,findCnt)
#     if len(listStr) <= findCnt:
#         return -1
#     return len(string)-len(listStr[-1])-len(subStr)


# def delete_err_name(li, correct_name, error_name):
#     for i in range(len(li)):
#         if li[i] == error_name:
#             li[i] = li[i][:findStr(li[i],correct_name,2):] 
#     return li


# # print(li[0].index('-'))
# print(delete_err_name(li, correct_name, error_name))


 
# a = "工程施工-差旅费-123123-444"
# sub = "-"

# N = 2      #查找第2次出现的位置
# # print(findStr(a,sub,N))

data=[
    [1,1,11,1,1],
    [2,2,2,2,2,2],
    [3,3,3,3,3,3]
]

import openpyxl

workbook=openpyxl.Workbook()

workbook_t=workbook.active


def write_2d_list_opxl(data,sheet):
    for i in range(len(data)):
        for j in range(len(data[i])):
            sheet.cell(i+1, j+1, data[i][j])

# write_2d_list_opxl()



workbook.save('test_3.xlsx')