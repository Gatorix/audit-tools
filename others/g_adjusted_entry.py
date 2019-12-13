import xlwt
import xlrd
import os
import msvcrt
import win32com.client as win32
import numpy
import glob
import inspect


def insert_blank_line(li, x, insert_list):
    '对二维列表相同行下面插入一个空行，li是需要操作的二维表，x是判断相同行的元素在行中的位置'
    n_li = []
    for i in range(1, len(li)):

        n_li.append(li[i-1])
        if li[i][x] != li[i-1][x]:
            n_li.append(insert_list)
    n_li.append(li[-1])
    return n_li


def write_header_line(sheet):
    header_line = [u'摘要', u'抵消编码', u'科目名称', u'借方金额', u'贷方金额', u'差异']
    for i in range(len(header_line)):
        sheet.write(0, i, header_line[i])


def write_list_nonformat(sheet, rows, cols, in_list):  # 写入list
    for i in range(rows):
        sheet.write(i+1, cols, in_list[i])


def write_list(sheet, rows, cols, in_list):  # 写入list
    style = xlwt.XFStyle()
    style.num_format_str = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
    for i in range(rows):
        sheet.write(i+1, cols, in_list[i], style)


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
            d_amount.append("0")
        else:
            c_amount.append("0")
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
    style = xlwt.XFStyle()
    style.num_format_str = '#,##0.00'
    for row in range(len(t_Schema)):
        for col in range(0, len(t_Schema[row])):
            sheet.write(row+start_row, col, t_Schema[row][col], style)


def get_col_value(sheet, col, start_row):  # 获取表格列值
    return sheet.col_values(col, start_row)


def exit_with_anykey():
    print("按任意键退出")
    ord(msvcrt.getch())
    os._exit(1)


def formatXLS(filepath):  # 转换为xlsx格式
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filepath)

    # FileFormat = 51 is for .xlsx extension
    # FileFormat = 56 is for .xls extension

    wb.SaveAs(filepath+"x", FileFormat=51)
    wb.Close()
    excel.Application.Quit()


INPUT_FILENAME = glob.glob(r"./海能达关联方交易及往来核对表*.xlsx")
if len(INPUT_FILENAME) == 0:
    print('错误：未找到包含"海能达关联方交易及往来核对表"的文件，确认后重试')
    exit_with_anykey()
try:
    print('检查往来核对表……')
    workbook_r = xlrd.open_workbook(INPUT_FILENAME[0])  # 打开工作簿
except FileNotFoundError as fnfe:
    print('错误：未找到文件，确认后重试')
    exit_with_anykey()
# print('Open workbook: ', workbook_r)
sheet_intercourse_elimination = workbook_r.sheet_by_name(
    '关联方往来抵消表')  # 获取'关联方往来抵消表'表
# print('Open sheet_intercourse_elimination at ', sheet_intercourse_elimination)

print("获取有效数据行数……")
nrows_Data = sheet_intercourse_elimination.nrows-1  # 获取有效数据行数，不含标题行

print("获取字段值……")
corp_sn = get_col_value(sheet_intercourse_elimination, 0, 1)  # 单位简称

o_corp_sn = get_col_value(sheet_intercourse_elimination, 1, 1)  # 对方单位简称

elimination_no = get_col_value(sheet_intercourse_elimination, 3, 1)  # 抵消编号

account_name = get_col_value(sheet_intercourse_elimination, 4, 1)  # 会计报表项目

d_c = trans_d_c(get_col_value(sheet_intercourse_elimination, 5, 1))  # 科目方向

amount = get_col_value(sheet_intercourse_elimination, 13, 1)  # 往来人民币余额
# print(type(amount[0]))
print("拼接摘要字段……")
abst = generate_abst(nrows_Data, corp_sn, o_corp_sn)  # 摘要
print("整理数据……")
o_list = []

for i in range(nrows_Data):  # 遍历各字段列表，按顺序合成新列表
    o_list.append(abst[i])
    if elimination_no[i] == "":
        o_list.append("None")
    else:
        o_list.append(int(elimination_no[i]))
    o_list.append(account_name[i])
    o_list.append(d_c[i])
    o_list.append(amount[i])

print('列表合并为矩阵……')
new_list = numpy.array(o_list).reshape(nrows_Data, 5)  # 列表合并为矩阵
print('拆分借贷方金额……')
c_amount, d_amount = split_dc(nrows_Data, new_list)  # 拆分借贷方金额
print('合并借贷方金额至矩阵……')
n_list = numpy.hstack((new_list, c_amount, d_amount))  # 合并借贷方金额至矩阵
print('删除无用字段……')
d_list = numpy.delete(n_list, [3, 4], axis=1).tolist()  # 删除无用字段
print('按抵消编码和借贷方排序……')
f_list = sorted(d_list, key=lambda x: (x[1], x[4]))  # 按抵消编码和借贷方排序
insert_list = ['', '', '', '0', '0', ]
i_list = insert_blank_line(f_list, 1, insert_list)

l_summary = numpy.array([x[0] for x in i_list],
                        dtype=str).tolist()  # 为不改变数值类型，曲线救国
l_elimination_no = numpy.array([x[1] for x in i_list], dtype=str).tolist()
l_account_name = numpy.array([x[2] for x in i_list], dtype=str).tolist()
l_c_amount = numpy.array([x[3] for x in i_list], dtype=float).tolist()
l_d_amount = numpy.array([x[4] for x in i_list], dtype=float).tolist()

len_i_list=len(l_summary)

val_list = [l_account_name, l_c_amount, l_d_amount]

print('创建调整分录表……')
workbook_w = xlwt.Workbook()  # 创建文件

sheet_intercourse_elimination_entry = workbook_w.add_sheet(
    '往来抵消分录', cell_overwrite_ok=True)  # 新建'往来抵消分录'，可复写

write_header_line(sheet_intercourse_elimination_entry)
print("写入调整分录……")

write_list_nonformat(sheet_intercourse_elimination_entry,
                     len_i_list, 0, l_summary)
write_list_nonformat(sheet_intercourse_elimination_entry,
                     len_i_list, 1, l_elimination_no)
for i in range(len(val_list)):
    write_list(sheet_intercourse_elimination_entry,
               len_i_list, i+2, val_list[i])


print('检查借贷金额……')
l_formula_str = []
for i in range(len_i_list):
    l_formula_str.append('SUMIF($B$2:$B$'+str(len_i_list+1)+',B'+str(i+2)+',$D$2:$D$' + str(len_i_list+1)+')-SUMIF($B$2:$B$'+str(len_i_list+1) +
                         ',B'+str(i+2)+',$E$2:$E$'+str(len_i_list+1)+')')

style = xlwt.XFStyle()
style.num_format_str = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

for x in range(len_i_list):
    sheet_intercourse_elimination_entry.write(
        x+1, 5, xlwt.Formula(l_formula_str[x]), style)


output_filename = "往来抵消分录.xls"
print("保存文件……")
try:
    workbook_w.save(output_filename)  # 保存文件
except PermissionError:
    print('错误：文件保存失败,关闭输出文件后重试')
    exit_with_anykey()
print("转换文件格式……")

current_path = os.path.abspath(__file__)
father_path = os.path.abspath(os.path.dirname(current_path) + os.path.sep + ".")

formatXLS(father_path+"\\往来抵消分录.xls")
os.remove(r"./往来抵消分录.xls")
print("完成")
exit_with_anykey()
