import xlwt
import xlrd
import os
import msvcrt
import win32com.client as win32
import numpy


def write_list(sheet, rows, cols, in_list):  # 写入list
    for i in range(rows):
        sheet.write(i+1, cols, in_list[i])


def generate_abst(nrows_Data, corp_sn, o_corp_sn):  # 拼接摘要字段
    abst = []
    for i in range(nrows_Data):
        abst.append(corp_sn[i]+"对"+o_corp_sn[i]+"往来抵消")
    return abst


def split_dc(nrows_Data, new_list):  # 借贷字段分列
    c_amount = []
    d_amount = []
    for i in range(nrows_Data):
        if new_list[i][3] == "借":
            c_amount.append(new_list[i][4])
            d_amount.append("")
        else:
            c_amount.append("")
            d_amount.append(new_list[i][4])
    return numpy.array(c_amount).reshape(nrows_Data, 1), numpy.array(d_amount).reshape(nrows_Data, 1)


def trans_d_c(dc_list):  # 转换借贷字段

    for i in range(nrows_Data):
        if dc_list[i] == "借":
            dc_list[i] = "贷"
        else:
            dc_list[i] = "借"
    return dc_list


def write_2d_list(sheet, t_Schema, start_row=0):  # 写入二维表
    for row in range(len(t_Schema)):
        for col in range(0, len(t_Schema[row])):
            sheet.write(row+start_row, col, t_Schema[row][col])


def get_col_value(sheet, col, start_row):  # 获取表格列值
    return sheet.col_values(col, start_row)


def exit_with_anykey():
    print("Press any key to exit...")
    ord(msvcrt.getch())
    os._exit(1)


def formatXLS(filepath):  # 转换为xls格式
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filepath)

    # FileFormat = 51 is for .xlsx extension
    # FileFormat = 56 is for .xls extension

    wb.SaveAs(filepath+"x", FileFormat=51)
    wb.Close()
    excel.Application.Quit()


INPUT_FILENAME = "D:\\audit-tools\\others\\海能达关联方交易及往来核对表 12.6.xlsx"

workbook_r = xlrd.open_workbook(INPUT_FILENAME)  # 打开工作簿

# print('Open workbook: ', workbook_r)
sheet_intercourse_elimination = workbook_r.sheet_by_name(
    '关联方往来抵消表')  # 获取'关联方往来抵消表'表
# print('Open sheet_intercourse_elimination at ', sheet_intercourse_elimination)


nrows_Data = sheet_intercourse_elimination.nrows-1  # 获取有效数据行数，不含标题行
# print(nrows_Data)

corp_sn = get_col_value(sheet_intercourse_elimination, 0, 1)  # 单位简称

o_corp_sn = get_col_value(sheet_intercourse_elimination, 1, 1)  # 对方单位简称


abst = generate_abst(nrows_Data, corp_sn, o_corp_sn)  # 摘要

elimination_no = get_col_value(sheet_intercourse_elimination, 3, 1)  # 抵消编号

account_name = get_col_value(sheet_intercourse_elimination, 4, 1)  # 会计报表项目

d_c = trans_d_c(get_col_value(sheet_intercourse_elimination, 5, 1))  # 科目方向

amount = get_col_value(sheet_intercourse_elimination, 13, 1)  # 往来人民币余额


o_list = []

for i in range(nrows_Data):  # 遍历各字段列表，按顺序合成新列表
    o_list.append(abst[i])
    if elimination_no[i] == "":
        o_list.append(0)
    else:
        o_list.append(int(elimination_no[i]))
    o_list.append(account_name[i])
    o_list.append(d_c[i])
    o_list.append(amount[i])

new_list = numpy.array(o_list).reshape(nrows_Data, 5)  # 列表合并为矩阵
# print(len(new_list))

c_amount, d_amount = split_dc(nrows_Data, new_list)  # 拆分借贷方金额

n_list = numpy.hstack((new_list, c_amount, d_amount))  # 合并借贷方金额至矩阵

d_list = numpy.delete(n_list, [3, 4], axis=1).tolist()  # 删除无用字段

f_list = sorted(d_list, key=lambda x: (x[1], x[4]))  # 按抵消编码和借贷方排序

# TODO 插入标题行

# TODO 修改数值格式

workbook_w = xlwt.Workbook()  # 创建文件

sheet_intercourse_elimination_entry = workbook_w.add_sheet(
    '往来抵消分录', cell_overwrite_ok=True)  # 新建'往来抵消分录'，可复写

write_2d_list(sheet_intercourse_elimination_entry, f_list)  # 写入值
output_filename = "往来抵消分录.xls"

workbook_w.save(output_filename)  # 保存文件
# formatXLS("D:\\audit-tools\\往来抵消分录.xls")
# os.remove("D:\\audit-tools\\往来抵消分录.xls")
exit_with_anykey()
